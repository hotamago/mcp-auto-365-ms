"""Synthetic Word packages only. No internal documents or network access."""

from __future__ import annotations

import copy
import io
import shutil
import subprocess
import urllib.error
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import tools as tools_mod
from common.errors import ConcurrentEditError, Mcp365Error, classify_http_error
from sharepoint import client as client_mod
from sharepoint import docx_comments as dc
from sharepoint import docx_review as review
from sharepoint.client import SharePointClient


def synthetic_docx(*, repeated=False, modern=False):
    xml = f'''<w:document xmlns:w="{dc.NS['w']}" xmlns:mc="{dc.NS['mc']}"
      xmlns:w14="{dc.NS['w14']}" xmlns:unused="urn:preserve-me" mc:Ignorable="unused w14">
      <w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">Before alpha </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t xml:space="preserve"> beta </w:t></w:r>
      <w:r><w:rPr><w:color w:val="123456"/></w:rPr><w:t>gamma after.</w:t></w:r></w:p>
      {'<w:p><w:r><w:t>alpha beta gamma</w:t></w:r></w:p>' if repeated else ''}
      <w:p><w:ins w:id="9" w:author="Other"><w:r><w:t>Tracked insertion</w:t></w:r></w:ins></w:p>
      <w:sectPr/></w:body></w:document>'''
    members = {
        '[Content_Types].xml': f'''<Types xmlns="{dc.NS['ct']}">
          <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
          <Default Extension="xml" ContentType="application/xml"/>
          <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
          </Types>'''.encode(),
        '_rels/.rels': f'<Relationships xmlns="{dc.NS["rel"]}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>'.encode(),
        'word/document.xml': xml.encode(),
        'word/styles.xml': f'<w:styles xmlns:w="{dc.NS["w"]}"/>'.encode(),
        'word/media/untouched.bin': bytes(range(256)) * 8,
    }
    if modern:
        for name in ('commentsIds', 'commentsExtensible'):
            prefix, root, _ = dc.PARTS[name]
            members[f'word/{name}.xml'] = dc.XmlPart.new(f'{prefix}:{root}').bytes()
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w') as z:
        z.comment = b'keep archive metadata'
        for i, (name, data) in enumerate(members.items()):
            info = zipfile.ZipInfo(name, (2025, 1, 2, 3, 4, 6))
            info.compress_type = zipfile.ZIP_STORED if i % 2 else zipfile.ZIP_DEFLATED
            info.comment = b'keep entry metadata'
            info.external_attr = 0o600 << 16
            z.writestr(info, data)
    return out.getvalue()


def parts(data):
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n) for n in z.namelist()}


def xml(data, name):
    return ET.fromstring(parts(data)[name])


def replace_parts(data, replacements):
    members = parts(data)
    members.update(replacements)
    out = io.BytesIO()
    with zipfile.ZipFile(out, 'w', zipfile.ZIP_DEFLATED) as archive:
        for name, value in members.items():
            if value is not None:
                archive.writestr(name, value)
    return out.getvalue()


def test_scoped_namespace_rebinding_and_qname_attributes_are_preserved():
    source = synthetic_docx()
    document = parts(source)['word/document.xml'].replace(
        b'<w:sectPr/>',
        b'<w:sectPr xmlns:unused="urn:rebound" mc:Ignorable="unused" w:rsidR="00000001"/>'
    )
    source = replace_parts(source, {'word/document.xml': document})
    result = dc.prepare(source, [{'anchor_text': 'alpha beta gamma', 'text': 'review'}], 'Reviewer')
    namespaces = [v for _, v in ET.iterparse(io.BytesIO(parts(result.data)['word/document.xml']), events=('start-ns',))]
    assert ('unused', 'urn:preserve-me') in namespaces
    assert ('unused', 'urn:rebound') in namespaces
    assert xml(result.data, 'word/document.xml').find('.//w:sectPr', dc.NS).get(dc.q('mc:Ignorable')) == 'unused'
    attrs = dc.XmlPart(b'<x title="a&#9;b&#10;c&#13;d"/>')
    assert ET.fromstring(attrs.bytes()).get('title') == 'a\tb\nc\rd'


def test_anchor_can_span_paragraphs_without_changing_document_text():
    source = synthetic_docx(repeated=True)
    result = dc.prepare(source, [{'anchor_text': 'gamma after. alpha beta', 'text': 'paragraph boundary'}], 'Reviewer')
    before = ''.join(t.text or '' for t in xml(source, 'word/document.xml').iter(dc.q('w:t')))
    after = ''.join(t.text or '' for t in xml(result.data, 'word/document.xml').iter(dc.q('w:t')))
    assert before == after
    dc.verify(result.data, result.records)


def test_reply_upgrades_legacy_parent_with_missing_para_id():
    first = dc.prepare(synthetic_docx(), [{'anchor_text': 'alpha', 'text': 'Legacy root'}], 'First Author')
    comments = xml(first.data, 'word/comments.xml')
    comments[0][0].attrib.pop(dc.q('w14:paraId'))
    source = replace_parts(first.data, {'word/comments.xml': ET.tostring(comments), 'word/commentsExtended.xml': None})
    result = dc.prepare(source, [{'reply_to': '0', 'text': 'reply'}], 'Reviewer')
    parent = xml(result.data, 'word/comments.xml')[0]
    assert parent.get(dc.q('w:author')) == 'First Author'
    assert dc._text(parent) == 'Legacy root'
    assert result.records[0]['parent_para_id'] == parent[-1].get(dc.q('w14:paraId'))


def test_complex_field_result_refused_without_mutation():
    source = synthetic_docx()
    doc = parts(source)['word/document.xml'].replace(
        b'<w:sectPr/>',
        b'<w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>DATE</w:instrText></w:r>'
        b'<w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>October 26</w:t></w:r>'
        b'<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:sectPr/>'
    )
    source = replace_parts(source, {'word/document.xml': doc})
    with pytest.raises(Mcp365Error, match='complex field'):
        dc.prepare(source, [{'anchor_text': 'October 26', 'text': 'review'}], 'Reviewer')


def test_formatted_anchor_split_across_three_runs_and_package_preserved():
    source = synthetic_docx()
    result = dc.prepare(source, [{'anchor_text': 'pha beta gam', 'text': 'Review <this> & that\nSecond paragraph'}], 'Reviewer Name')
    old, new = parts(source), parts(result.data)
    for name in old.keys() - {'word/document.xml', '[Content_Types].xml'}:
        assert old[name] == new[name]
    with zipfile.ZipFile(io.BytesIO(source)) as a, zipfile.ZipFile(io.BytesIO(result.data)) as b:
        assert b.namelist()[:len(a.namelist())] == a.namelist()
        assert b.comment == a.comment
        for first, second in zip(a.infolist(), b.infolist(), strict=False):
            assert (first.compress_type, first.date_time, first.comment, first.external_attr) == (second.compress_type, second.date_time, second.comment, second.external_attr)
    for name, data in new.items():
        if name.endswith('.xml') or name.endswith('.rels'):
            ET.fromstring(data)
    assert b'xmlns:unused="urn:preserve-me"' in new['word/document.xml']
    assert b'mc:Ignorable="unused w14"' in new['word/document.xml']
    doc = xml(result.data, 'word/document.xml')
    para = doc.find('.//w:p', dc.NS)
    selected, in_range = [], False
    for node in para:
        if node.tag == dc.q('w:commentRangeStart'):
            in_range = True
        elif node.tag == dc.q('w:commentRangeEnd'):
            in_range = False
        elif in_range:
            selected.extend(t.text or '' for t in node.iter(dc.q('w:t')))
    assert ''.join(selected) == 'pha  beta gam'
    runs = [r for r in para.findall('w:r', dc.NS) if r.find('w:t', dc.NS) is not None]
    assert [(r.find('w:rPr', dc.NS)[0].tag.rsplit('}', 1)[-1], r.find('w:t', dc.NS).text) for r in runs] == [
        ('b', 'Before al'), ('b', 'pha '), ('i', ' beta '), ('color', 'gam'), ('color', 'ma after.')]
    before_revision = xml(source, 'word/document.xml').find('.//w:ins', dc.NS)
    assert ET.tostring(before_revision) == ET.tostring(doc.find('.//w:ins', dc.NS))
    comment = xml(result.data, 'word/comments.xml')[0]
    assert comment.get(dc.q('w:initials')) == 'RN'
    assert not comment.get(dc.q('w:date')).endswith('Z')
    assert len(comment.findall('w:p', dc.NS)) == 2
    assert len({p.get(dc.q('w14:paraId')) for p in comment}) == 2
    for name in ('comments', 'commentsExtended'):
        assert any(e.get('PartName') == f'/word/{name}.xml' for e in xml(result.data, '[Content_Types].xml'))
        assert any(e.get('Target') == f'{name}.xml' for e in xml(result.data, 'word/_rels/document.xml.rels'))


def test_existing_thread_reply_last_paragraph_durable_ids_and_mirrored_ranges():
    first = dc.prepare(synthetic_docx(modern=True), [{'anchor_text': 'alpha beta gamma', 'text': 'Root\nLast paragraph'}], 'First Author')
    second = dc.prepare(first.data, [{'reply_to': first.records[0]['id'], 'text': 'Existing reply'}], 'Second Author')
    old_comments = copy.deepcopy(xml(second.data, 'word/comments.xml'))
    result = dc.prepare(second.data, [{'reply_to': 'First Author: Root', 'text': 'New reply'}], 'Third Author')
    new_comments = xml(result.data, 'word/comments.xml')
    assert len({c.get(dc.q('w:id')) for c in new_comments}) == 3
    assert [ET.tostring(c) for c in new_comments[:2]] == [ET.tostring(c) for c in old_comments]
    root_para = first.records[0]['para_id']
    assert result.records[0]['parent_para_id'] == root_para
    entries = xml(result.data, 'word/commentsExtended.xml')
    assert [e.get(dc.q('w15:paraIdParent')) for e in entries] == [None, root_para, root_para]
    identifiers = xml(result.data, 'word/commentsIds.xml')
    durable = [e.get(dc.q('w16cid:durableId')) for e in identifiers]
    assert len(set(durable)) == 3
    extensible = xml(result.data, 'word/commentsExtensible.xml')
    assert [e.get(dc.q('w16cex:durableId')) for e in extensible] == durable
    assert all(e.get(dc.q('w16cex:dateUtc')).endswith('Z') for e in extensible)
    doc = xml(result.data, 'word/document.xml')
    started, ended, ranges, text = {}, {}, {}, ''
    for e in doc.iter():
        if e.tag == dc.q('w:t'):
            text += e.text or ''
        if e.tag == dc.q('w:commentRangeStart'):
            started[e.get(dc.q('w:id'))] = len(text)
        if e.tag == dc.q('w:commentRangeEnd'):
            ended[e.get(dc.q('w:id'))] = len(text)
    for cid in started:
        ranges[cid] = (started[cid], ended[cid])
    assert len(set(ranges.values())) == 1
    dc.verify(result.data, first.records + second.records + result.records)


@pytest.mark.parametrize('spec, message', [
    ({'anchor_text': 'missing', 'text': 'x'}, 'not found'),
    ({'anchor_text': 'alpha beta gamma', 'text': 'x'}, 'Ambiguous'),
    ({'anchor_text': 'alpha beta gamma', 'occurrence': 3, 'text': 'x'}, 'only 2'),
    ({'anchor_text': 'alpha', 'occurrence': 0, 'text': 'x'}, 'positive'),
    ({'reply_to': '99', 'text': 'x'}, 'not found'),
    ({'anchor_text': 'Tracked insertion', 'text': 'x'}, 'tracked revision'),
    ({'anchor_text': 'alpha', 'reply_to': '0', 'text': 'x'}, 'exactly one'),
])
def test_invalid_batch_aborts_without_changing_input(spec, message):
    source = synthetic_docx(repeated=True)
    original = bytes(source)
    with pytest.raises(Mcp365Error, match=message):
        dc.prepare(source, [{'anchor_text': 'Before', 'text': 'valid first'}, spec], 'Reviewer')
    assert source == original


def test_explicit_occurrence_and_overlapping_batch():
    source = synthetic_docx(repeated=True)
    result = dc.prepare(source, [
        {'anchor_text': 'alpha beta gamma', 'occurrence': 2, 'text': 'second'},
        {'anchor_text': 'beta', 'occurrence': 1, 'text': 'overlapping'},
        {'anchor_text': 'alpha beta gamma', 'occurrence': 1, 'text': 'first'},
    ], 'Reviewer')
    dc.verify(result.data, result.records)


@pytest.fixture
def remote(monkeypatch):
    client = SharePointClient()
    state = {'data': synthetic_docx(), 'etag': '"file-guid,1"', 'version': '1.0', 'calls': [], 'conflict': False, 'bad_verify': False, 'change_during_download': False}
    site = 'https://example-my.sharepoint.com/personal/reviewer'
    client._drive_sites['drive'] = site
    client._drive_cache['web:drive'] = site + '/Documents'
    monkeypatch.setattr(client, 'resolve_file', lambda _: ('drive', {'name': 'review.docx', 'id': 'item', 'parentReference': {'path': '/drive/root:/Claude-test'}}))
    monkeypatch.setattr(client, '_cookie_headers', lambda **_: {'Cookie': 'not-a-real-secret'})
    monkeypatch.setattr(client, '_get_form_digest', lambda target: 'synthetic-digest' if target == site else pytest.fail('wrong digest site'))
    monkeypatch.setattr(client, 'get_token', lambda **_: pytest.fail('Graph must not be called'))

    def http(url, *, headers, method='GET', data=None, **kwargs):
        state['calls'].append((url, method, dict(headers), data, kwargs))
        if url.endswith('/$value') or method == 'PUT':
            if method == 'PUT':
                assert url == site + '/Documents/Claude-test/review.docx'
                assert 'X-HTTP-Method' not in headers
                assert headers['If-Match'] and headers['If-Match'] != '*'
                assert headers['X-RequestDigest'] == 'synthetic-digest'
                assert kwargs['max_retries'] == 0
                if state['conflict'] or headers['If-Match'] != state['etag']:
                    raise classify_http_error(urllib.error.HTTPError(url, 412, 'Precondition failed', {}, None), 'review')
                state['data'] = data if not state['bad_verify'] else synthetic_docx()
                state['etag'], state['version'] = '"file-guid,2"', '2.0'
            if method == 'GET' and state['change_during_download']:
                state['etag'] = '"file-guid,2"'
            return 200, state['data'], {'ETag': state['etag']}
        if 'currentuser' in url:
            return {'d': {'Title': 'Session User'}}
        return {'d': {'ETag': state['etag'], 'UIVersionLabel': state['version']}}

    monkeypatch.setattr(review, 'request', http)
    monkeypatch.setattr(review, 'request_json', http)
    return client, state


def writes(state):
    return [c for c in state['calls'] if c[1] != 'GET']


def test_dry_run_resolves_and_shows_exact_preview_without_upload(remote):
    client, state = remote
    text = review.add_comments(client, 'https://file', [{'anchor_text': 'alpha beta gamma', 'text': 'Exact review'}], False)
    assert 'Exact review' in text and 'Session User' in text and 'Before alpha beta gamma after.' in text
    assert '"file-guid,1"' in text and '1.0' in text and 'CHƯA GỬI' in text
    assert not writes(state)


def test_conditional_upload_and_new_version_verified(remote):
    client, state = remote
    text = review.add_comments(client, 'https://file', [{'anchor_text': 'alpha beta gamma', 'text': 'Exact review'}], True)
    assert 'verified 1' in text and 'Version: 2.0' in text
    assert len(writes(state)) == 1
    assert len([c for c in state['calls'] if c[0].endswith('/$value') and c[1] == 'GET']) == 2
    assert b'Exact review' in parts(state['data'])['word/comments.xml']


def test_conditional_write_uses_webdav_never_rest_value_which_ignores_if_match(remote):
    client, state = remote
    review.add_comments(client, 'https://file', [{'anchor_text': 'alpha', 'text': 'review'}], True)
    write = writes(state)[0]
    assert write[1] == 'PUT'
    assert '/_api/' not in write[0]
    assert write[0].endswith('/Documents/Claude-test/review.docx')
    assert write[2]['If-Match'] == '"file-guid,1"'
    assert write[4]['max_retries'] == 0


def test_412_never_retries_or_unconditionally_overwrites(remote):
    client, state = remote
    state['conflict'] = True
    original = state['data']
    with pytest.raises(ConcurrentEditError, match='File changed since download'):
        review.add_comments(client, 'https://file', [{'anchor_text': 'alpha', 'text': 'review'}], True)
    assert len(writes(state)) == 1
    assert writes(state)[0][2]['If-Match'] != '*'
    assert state['data'] == original


def test_stale_snapshot_upload_refuses_even_if_remote_changed_in_another_call(remote):
    client, state = remote
    old = review.snapshot(client, 'https://file')
    state['data'] = b'changed by another editor'
    state['etag'] = '"file-guid,2"'
    with pytest.raises(ConcurrentEditError):
        review.upload(client, old, old.content)
    assert state['data'] == b'changed by another editor'
    assert len(writes(state)) == 1


def test_snapshot_changed_while_downloading_aborts_before_upload(remote):
    client, state = remote
    state['change_during_download'] = True
    with pytest.raises(ConcurrentEditError, match='File changed since download'):
        review.add_comments(client, 'https://file', [{'anchor_text': 'alpha', 'text': 'review'}], True)
    assert not writes(state)


@pytest.mark.parametrize('etag', ['', '*', 'W/"weak"'])
def test_missing_or_unsafe_etag_refuses_upload(remote, etag):
    client, state = remote
    state['etag'] = etag
    with pytest.raises(Mcp365Error, match='strong file ETag'):
        review.add_comments(client, 'https://file', [{'anchor_text': 'alpha', 'text': 'review'}], True)
    assert not writes(state)


def test_confirmed_call_reads_a_fresh_etag_after_preview(remote):
    client, state = remote
    comments = [{'anchor_text': 'alpha', 'text': 'review'}]
    review.add_comments(client, 'https://file', comments, False)
    state['etag'] = '"file-guid,9"'
    review.add_comments(client, 'https://file', comments, True)
    assert writes(state)[0][2]['If-Match'] == '"file-guid,9"'


def test_verification_failure_reports_upload_already_happened(remote):
    client, state = remote
    state['bad_verify'] = True
    with pytest.raises(Mcp365Error, match='upload succeeded, but verification failed'):
        review.add_comments(client, 'https://file', [{'anchor_text': 'alpha', 'text': 'review'}], True)
    assert len(writes(state)) == 1


def test_cookie_only_never_falls_back_on_failed_resolution(monkeypatch):
    client = SharePointClient()
    monkeypatch.setattr(client, '_cookie_request', lambda *a, **k: (_ for _ in ()).throw(Mcp365Error('cookie failed')))
    monkeypatch.setattr(client, '_graph_json', lambda *a, **k: pytest.fail('Graph fallback'))
    with client.cookie_only(), pytest.raises(Mcp365Error, match='cookie failed'):
        client.resolve_file('https://example-my.sharepoint.com/personal/reviewer/Documents/file.docx')
    assert client_mod._COOKIE_ONLY.get() is False


@pytest.mark.anyio
async def test_mcp_dry_run_and_actionable_failure(remote, monkeypatch):
    client, state = remote
    monkeypatch.setattr(tools_mod, 'sp', lambda: client)
    server = MCPServer('docx-test')
    tools_mod.register_all(server)
    result = await server.call_tool('add_docx_comments', {'url_or_guid': 'https://file', 'comments': [{'anchor_text': 'alpha', 'text': 'MCP preview'}], 'is_user_confirm': False})
    assert 'MCP preview' in str(result) and not writes(state)
    with pytest.raises(ToolError, match='Nothing uploaded'):
        await server.call_tool('add_docx_comments', {'url_or_guid': 'https://file', 'comments': [{'anchor_text': 'missing', 'text': 'bad'}], 'is_user_confirm': True})
    assert not writes(state)


@pytest.mark.skipif(shutil.which('soffice') is None, reason='LibreOffice is not installed')
def test_libreoffice_opens_comments_and_reply(tmp_path):
    root = dc.prepare(synthetic_docx(modern=True), [{'anchor_text': 'alpha beta gamma', 'text': 'Root review'}], 'Reviewer')
    result = dc.prepare(root.data, [{'reply_to': root.records[0]['id'], 'text': 'Reply review'}], 'Reviewer')
    local = tmp_path / 'review.docx'
    local.write_bytes(result.data)
    converted = subprocess.run([
        shutil.which('soffice'), f'-env:UserInstallation={Path(tmp_path / "profile").as_uri()}',
        '--headless', '--convert-to', 'pdf', '--outdir', str(tmp_path), str(local),
    ], capture_output=True, text=True, timeout=45)
    assert converted.returncode == 0, converted.stdout + converted.stderr
    assert (tmp_path / 'review.pdf').read_bytes().startswith(b'%PDF')
