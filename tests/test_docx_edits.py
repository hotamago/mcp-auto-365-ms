"""Synthetic OOXML and mocked conditional round trips; never contact a tenant."""

from __future__ import annotations

import copy
import io
import shutil
import subprocess
import urllib.error
import zipfile
from datetime import datetime

import pytest
import test_docx_comments as fixtures
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError
from PIL import Image
from test_docx_comments import parts, replace_parts, synthetic_docx, writes, xml

import tools as tools_mod
from common.errors import ConcurrentEditError, Mcp365Error, classify_http_error
from sharepoint import docx_comments as dc
from sharepoint import docx_edits as de
from sharepoint import docx_review as review


@pytest.fixture
def remote(monkeypatch):
    return fixtures.remote.__wrapped__(monkeypatch)


def raster(size=(100, 50), fmt='PNG', color='red'):
    out = io.BytesIO()
    Image.new('RGB', size, color).save(out, format=fmt)
    return out.getvalue()


def image_docx(*, shared=False, alias=False, header=False, root_media=False, caption_before=False):
    source = synthetic_docx()
    drawing = f'''<w:r><w:drawing xmlns:wp="{de.WP}" xmlns:a="{de.A}" xmlns:pic="{de.PIC}" xmlns:r="{de.R}">
      <wp:inline><wp:extent cx="1000000" cy="500000"/><wp:docPr id="1" name="Diagram A"/>
      <a:graphic><a:graphicData uri="{de.PIC}"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="diagram.png"/><pic:cNvPicPr/></pic:nvPicPr>
      <pic:blipFill><a:blip r:embed="rId7"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/>
      <a:ext cx="1000000" cy="500000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>'''
    caption = '<w:p><w:r><w:t>Figure A — unique caption</w:t></w:r></w:p>'
    block = f'<w:p>{drawing}</w:p>'
    block = caption + block if caption_before else block + caption
    if shared or alias:
        other = drawing.replace('rId7', 'rId8') if alias else drawing
        block += '<w:p><w:r><w:t>separator</w:t></w:r></w:p><w:p>' + other + '</w:p>'
    doc = parts(source)[de.DOC].replace(b'<w:sectPr/>', (block + '<w:sectPr/>').encode())
    target = '/media/original.png' if root_media else 'media/original.png'
    media = 'media/original.png' if root_media else 'word/media/original.png'
    rels = f'<Relationships xmlns="{dc.NS["rel"]}"><Relationship Id="rId7" Type="{de.R}/image" Target="{target}"/>'
    if alias:
        rels += f'<Relationship Id="rId8" Type="{de.R}/image" Target="{target}"/>'
    replacements = {de.DOC: doc, de.RELS: (rels + '</Relationships>').encode(), media: raster()}
    if header:
        replacements['word/_rels/header1.xml.rels'] = f'<Relationships xmlns="{dc.NS["rel"]}"><Relationship Id="rId1" Type="{de.R}/image" Target="{target}"/></Relationships>'.encode()
        replacements['word/header1.xml'] = b'<untouched/>'
    return replace_parts(source, replacements)


def test_cross_run_replacement_accept_reject_and_formatting():
    source = synthetic_docx()
    result = de.prepare_text(source, [{'find': 'alpha  beta gamma', 'replace': 'New words'}], 'Session User')
    document = xml(result.data, de.DOC)
    paragraph = document.find('.//w:p', dc.NS)
    assert ''.join(t.text or '' for t in paragraph.iter(dc.q('w:t'))) == 'Before New words after.'
    assert ''.join(t.text or '' for t in paragraph.iter(dc.q('w:delText'))) == 'alpha  beta gamma'
    inserted = paragraph.find('w:ins', dc.NS)
    assert inserted.find('w:r/w:rPr/w:b', dc.NS) is not None
    for node in paragraph.findall('w:del', dc.NS) + [inserted]:
        assert node.get(dc.q('w:author')) == 'Session User'
        assert datetime.fromisoformat(node.get(dc.q('w:date'))).tzinfo is None
    ids = [n.get(dc.q('w:id')) for n in document.iter() if n.tag in {dc.q('w:ins'), dc.q('w:del')}]
    assert len(ids) == len(set(ids))
    # Rejecting the revisions restores the exact text and formatting runs.
    rejected = copy.deepcopy(paragraph)
    for child in list(rejected):
        if child.tag == dc.q('w:ins'):
            rejected.remove(child)
        elif child.tag == dc.q('w:del'):
            at = list(rejected).index(child)
            rejected.remove(child)
            for run in child:
                for t in run.iter(dc.q('w:delText')):
                    t.tag = dc.q('w:t')
                rejected.insert(at, run)
                at += 1
    assert ''.join(t.text or '' for t in rejected.iter(dc.q('w:t'))) == 'Before alpha  beta gamma after.'
    before, after = parts(source), parts(result.data)
    assert all(after[n] == v for n, v in before.items() if n != de.DOC)
    with zipfile.ZipFile(io.BytesIO(source)) as a, zipfile.ZipFile(io.BytesIO(result.data)) as b:
        assert a.comment == b.comment
        for i, j in zip(a.infolist(), b.infolist(), strict=True):
            assert (i.filename, i.date_time, i.compress_type, i.comment, i.external_attr) == (j.filename, j.date_time, j.compress_type, j.comment, j.external_attr)


def test_insert_whole_paragraphs_tracks_marks_and_copies_style():
    source = synthetic_docx()
    doc = parts(source)[de.DOC].replace(b'<w:p>', b'<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr>', 1)
    source = replace_parts(source, {de.DOC: doc})
    result = de.prepare_text(source, [{'insert_after': 'Before alpha', 'text': 'New subsection\nNew body'}], 'Reviewer')
    paragraphs = xml(result.data, de.DOC).findall('w:body/w:p', dc.NS)
    for paragraph, text in zip(paragraphs[1:3], ['New subsection', 'New body'], strict=True):
        assert paragraph.find('w:pPr/w:pStyle', dc.NS).get(dc.q('w:val')) == 'BodyText'
        assert paragraph.find('w:pPr/w:rPr/w:ins', dc.NS) is not None
        assert paragraph.find('w:ins/w:r/w:t', dc.NS).text == text
    assert paragraphs[3].find('w:ins', dc.NS).get(dc.q('w:author')) == 'Other'


@pytest.mark.parametrize('edits,match', [
    ([{'find': 'missing', 'replace': 'x'}], 'not found'),
    ([{'find': 'alpha', 'replace': 'x', 'occurrence': 0}], 'positive'),
    ([{'find': 'alpha', 'replace': 'x', 'occurrence': True}], 'positive'),
    ([{'find': 'alpha', 'replace': 'x', 'occurrence': 5}], 'found only'),
    ([{'find': 'alpha', 'replace': 'x'}, {'find': 'pha', 'replace': 'y'}], 'Overlapping'),
    ([{'find': 'Tracked insertion', 'replace': 'x'}], 'tracked revisions'),
    ([{'find': 'alpha', 'replace': 'x\ny'}], 'one paragraph'),
    ([{'find': 'alpha', 'replace': '\x00'}], 'XML-safe'),
    ([{'find': 'alpha', 'replace': 'x', 'extra': 'ignored?'}], 'Use'),
    ([], 'non-empty'),
])
def test_invalid_batch_aborts(edits, match):
    source = synthetic_docx()
    with pytest.raises(Mcp365Error, match=match):
        de.prepare_text(source, edits, 'Reviewer')


def test_ambiguous_occurrence_deletion_and_multiple_edits():
    source = synthetic_docx(repeated=True)
    with pytest.raises(Mcp365Error, match='Ambiguous'):
        de.prepare_text(source, [{'find': 'alpha', 'replace': 'x'}], 'Reviewer')
    result = de.prepare_text(source, [
        {'find': 'alpha', 'replace': '', 'occurrence': 2},
        {'find': 'Before', 'replace': 'Start'},
        {'find': 'after.', 'replace': 'end.'},
    ], 'Reviewer')
    paragraphs = xml(result.data, de.DOC).findall('w:body/w:p', dc.NS)
    assert ''.join(t.text or '' for t in paragraphs[0].iter(dc.q('w:t'))) == 'Start alpha  beta gamma end.'
    assert ''.join(t.text or '' for t in paragraphs[1].iter(dc.q('w:t'))) == ' beta gamma'


@pytest.mark.parametrize('markup', [
    '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:t>FIELD</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>',
    '<w:fldSimple w:instr="REF"><w:r><w:t>FIELD</w:t></w:r></w:fldSimple>',
    '<w:del w:id="55"><w:r><w:delText>old</w:delText></w:r></w:del>',
])
def test_range_crossing_field_or_revision_refused(markup):
    source = synthetic_docx()
    doc = parts(source)[de.DOC].replace(b'<w:r><w:rPr><w:i/>', markup.encode() + b'<w:r><w:rPr><w:i/>')
    with pytest.raises(Mcp365Error, match='tracked revisions'):
        de.prepare_text(replace_parts(source, {de.DOC: doc}), [{'find': 'Before', 'replace': 'x'}], 'Reviewer')


def test_existing_comments_and_bookmarks_survive():
    source = dc.prepare(synthetic_docx(), [{'anchor_text': 'alpha beta gamma', 'text': 'Keep comment'}], 'Old reviewer').data
    result = de.prepare_text(source, [{'find': 'alpha  beta gamma', 'replace': 'New'}], 'Reviewer')
    assert parts(result.data)['word/comments.xml'] == parts(source)['word/comments.xml']
    for tag in ('commentRangeStart', 'commentRangeEnd', 'commentReference'):
        assert len(list(xml(result.data, de.DOC).iter(dc.q('w:' + tag)))) == 1


def test_literal_tabs_are_tracked_with_the_surrounding_text():
    source = synthetic_docx()
    doc = parts(source)[de.DOC].replace(b'<w:t>gamma after.</w:t>', b'<w:tab/><w:t>gamma after.</w:t>')
    source = replace_parts(source, {de.DOC: doc})
    result = de.prepare_text(source, [{'find': 'beta \tgamma', 'replace': 'New'}], 'Reviewer')
    paragraph = xml(result.data, de.DOC).find('.//w:p', dc.NS)
    assert paragraph.find('w:del/w:r/w:tab', dc.NS) is not None
    assert all(node.find('w:tab', dc.NS) is None for node in paragraph.findall('w:r', dc.NS))


def test_complex_field_spanning_paragraphs_refused():
    source = synthetic_docx()
    doc = parts(source)[de.DOC].replace(b'<w:body>', b'<w:body><w:p><w:r><w:fldChar w:fldCharType="begin"/></w:r></w:p>')
    doc = doc.replace(b'<w:sectPr/>', b'<w:p><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p><w:sectPr/>')
    with pytest.raises(Mcp365Error, match='complex field result'):
        de.prepare_text(replace_parts(source, {de.DOC: doc}), [{'find': 'Before', 'replace': 'New'}], 'Reviewer')


@pytest.mark.parametrize('caption_before', [False, True])
@pytest.mark.parametrize('root_media', [False, True])
def test_image_caption_replaces_only_one_media_and_preserves_extent(caption_before, root_media):
    source = image_docx(caption_before=caption_before, root_media=root_media)
    replacement = raster(color='blue')
    result, current = de.prepare_image(source, {'caption_text': 'Figure A'}, replacement)
    assert current == raster()
    assert result.preview[0]['extent_before'] == result.preview[0]['extent_after']
    assert result.preview[0]['relationship_id_before'] == result.preview[0]['relationship_id_after'] == 'rId7'
    assert result.preview[0]['shared_media'] is False
    changed = {n for n in parts(source) if parts(source)[n] != parts(result.data)[n]}
    assert changed <= {de.DOC, result.preview[0]['current_media']}
    assert parts(result.data)[result.preview[0]['output_media']] == replacement


@pytest.mark.parametrize('options,new_rid', [({'shared': True}, True), ({'alias': True}, False), ({'header': True}, False)])
def test_shared_media_isolated_without_changing_other_users(options, new_rid):
    source = image_docx(**options)
    result, _ = de.prepare_image(source, {'caption_text': 'Figure A'}, raster(color='blue'))
    p = result.preview[0]
    assert p['shared_media'] is True
    assert p['output_media'] != p['current_media']
    assert parts(result.data)[p['current_media']] == parts(source)[p['current_media']]
    assert (p['relationship_id_after'] != p['relationship_id_before']) == new_rid
    if options.get('shared'):
        blips = list(xml(result.data, de.DOC).iter(f'{{{de.A}}}blip'))
        assert blips[0].get(f'{{{de.R}}}embed') != blips[1].get(f'{{{de.R}}}embed')


def test_format_change_allocates_correct_typed_part():
    result, _ = de.prepare_image(image_docx(), {'image_name': 'Diagram A'}, raster(fmt='JPEG'))
    p = result.preview[0]
    assert p['output_media'].endswith('.jpg')
    assert p['relationship_id_after'] == 'rId7'
    types = xml(result.data, '[Content_Types].xml')
    assert any(n.get('PartName') == '/' + p['output_media'] and n.get('ContentType') == 'image/jpeg' for n in types)


@pytest.mark.parametrize('size,expected', [((100, 100), 1000000), ((1000, 501), 500000)])
def test_image_ratio_threshold_and_width_preserved(size, expected):
    result, _ = de.prepare_image(image_docx(), {'image_name': 'diagram.png'}, raster(size=size))
    extent = result.preview[0]['extent_after']
    assert extent == {'cx': 1000000, 'cy': expected}
    for node in xml(result.data, de.DOC).iter(f'{{{de.A}}}ext'):
        assert int(node.get('cy')) == expected


@pytest.mark.parametrize('locate', [{}, {'caption_text': 'absent'}, {'image_name': 'Diagram A', 'caption_text': 'Figure A'}])
def test_invalid_image_locator_refused(locate):
    with pytest.raises(Mcp365Error):
        de.prepare_image(image_docx(), locate, raster())


def test_ambiguous_image_and_invalid_image_refused():
    with pytest.raises(Mcp365Error, match='2 drawings'):
        de.prepare_image(image_docx(shared=True), {'image_name': 'Diagram A'}, raster())
    for data in (b'not a PNG', raster()[:40], raster(fmt='GIF')):
        with pytest.raises(Mcp365Error):
            de.prepare_image(image_docx(), {'caption_text': 'Figure A'}, data)


def test_verifiers_detect_revision_or_media_tampering():
    source = synthetic_docx()
    result = de.prepare_text(source, [{'find': 'alpha', 'replace': 'changed'}], 'Reviewer')
    with pytest.raises(Mcp365Error, match='failed verification'):
        de.verify_text(source, result.records)
    source = image_docx()
    result, _ = de.prepare_image(source, {'caption_text': 'Figure A'}, raster(color='blue'))
    with pytest.raises(Mcp365Error, match='failed verification'):
        de.verify_image(source, result.records)


@pytest.mark.parametrize('value', [False, None, 'true', 1])
def test_text_preview_never_writes_and_derives_author(remote, value):
    client, state = remote
    result = review.edit_text(client, 'https://file', [{'find': 'alpha', 'replace': 'New'}], value)
    assert 'CHƯA GỬI' in result and 'Session User' in result and 'New' in result
    assert not writes(state)


@pytest.mark.parametrize('kind', ['text', 'image'])
def test_confirmed_round_trip_and_verification_failure(remote, tmp_path, kind):
    client, state = remote
    if kind == 'text':
        def call():
            return review.edit_text(client, 'https://file', [{'find': 'alpha', 'replace': 'New'}], True)
    else:
        state['data'] = image_docx()
        path = tmp_path / 'new.png'
        path.write_bytes(raster(color='blue'))
        def call():
            return review.replace_image(client, 'https://file', {'caption_text': 'Figure A'}, str(path), True)
    assert 're-downloaded/verified' in call()
    assert len(writes(state)) == 1
    assert writes(state)[0][2]['If-Match'] == '"file-guid,1"'
    assert writes(state)[0][4]['max_retries'] == 0
    state['data'] = synthetic_docx() if kind == 'text' else image_docx()
    state['bad_verify'] = True
    with pytest.raises(Mcp365Error, match='upload succeeded, but verification failed'):
        call()


def test_image_preview_extracts_original_and_does_not_write(remote, tmp_path):
    client, state = remote
    state['data'] = image_docx()
    path = tmp_path / 'new.png'
    path.write_bytes(raster(color='blue'))
    text = review.replace_image(client, 'https://file', {'caption_text': 'Figure A'}, str(path), False)
    import json
    from pathlib import Path
    preview = json.loads(text[text.index('[\n'):text.rindex(']') + 1])[0]
    assert Path(preview['current_image_path']).read_bytes() == raster()
    assert preview['current_image']['width'] == 100 and preview['current_image']['height'] == 50
    assert not writes(state)


@pytest.mark.parametrize('code', [409, 412, 423])
@pytest.mark.parametrize('kind', ['text', 'image'])
def test_conflict_is_exactly_one_conditional_attempt(remote, monkeypatch, tmp_path, code, kind):
    client, state = remote
    original_request = review.request
    def request(url, **kwargs):
        if kwargs.get('method') == 'PUT':
            state['calls'].append((url, 'PUT', kwargs['headers'], None, kwargs))
            raise classify_http_error(urllib.error.HTTPError(url, code, 'locked or changed', {}, None), 'DOCX')
        return original_request(url, **kwargs)
    monkeypatch.setattr(review, 'request', request)
    if kind == 'image':
        state['data'] = image_docx()
    original = state['data']
    with pytest.raises(ConcurrentEditError, match='file locked or changed; retry later'):
        if kind == 'text':
            review.edit_text(client, 'https://file', [{'find': 'alpha', 'replace': 'New'}], True)
        else:
            path = tmp_path / 'new.png'
            path.write_bytes(raster(color='blue'))
            review.replace_image(client, 'https://file', {'caption_text': 'Figure A'}, str(path), True)
    assert state['data'] == original
    assert len(writes(state)) == 1
    assert writes(state)[0][2]['If-Match'] == '"file-guid,1"'


@pytest.mark.anyio
async def test_new_tools_registered_and_callable_without_mutation(remote, monkeypatch, tmp_path, all_tools):
    client, state = remote
    monkeypatch.setattr(tools_mod, 'sp', lambda: client)
    server = MCPServer('test')
    tools_mod.register_all(server)
    result = await server.call_tool('edit_docx_text', {'url': 'https://file', 'edits': [{'find': 'alpha', 'replace': 'New'}], 'is_user_confirm': False})
    assert 'CHƯA GỬI' in str(result)
    state['data'] = image_docx()
    path = tmp_path / 'new.png'
    path.write_bytes(raster())
    result = await server.call_tool('replace_docx_image', {'url': 'https://file', 'locate': {'caption_text': 'Figure A'}, 'new_image_path': str(path), 'is_user_confirm': False})
    assert 'current_image_path' in str(result) and not writes(state)


@pytest.mark.anyio
@pytest.mark.parametrize('name', ['edit_docx_text', 'replace_docx_image'])
@pytest.mark.parametrize('value', ['true', 'false', 1, None])
async def test_mcp_does_not_coerce_nonboolean_confirmation(monkeypatch, name, value, all_tools):
    monkeypatch.setattr(tools_mod, 'sp', lambda: pytest.fail('nonboolean confirmation must fail before client access'))
    server = MCPServer('strict-confirmation')
    tools_mod.register_all(server)
    args = {'url': 'https://file', 'is_user_confirm': value}
    if name == 'edit_docx_text':
        args['edits'] = [{'find': 'x', 'replace': 'y'}]
    else:
        args.update(locate={'caption_text': 'Figure A'}, new_image_path='/tmp/example.png')
    with pytest.raises(ToolError, match='valid boolean'):
        await server.call_tool(name, args)


@pytest.mark.skipif(shutil.which('soffice') is None, reason='LibreOffice is not installed')
def test_libreoffice_round_trip_retains_real_revisions_and_shared_images(tmp_path):
    source = image_docx(shared=True)
    revised = de.prepare_text(source, [
        {'find': 'alpha  beta gamma', 'replace': 'Replacement words'},
        {'insert_after': 'separator', 'text': 'Inserted section\nInserted body'},
    ], 'Reviewer')
    edited, _ = de.prepare_image(revised.data, {'caption_text': 'Figure A'}, raster(size=(100, 100), color='blue'))
    local = tmp_path / 'review.docx'
    local.write_bytes(edited.data)
    out = tmp_path / 'converted'
    out.mkdir()
    converted = subprocess.run([
        shutil.which('soffice'), f'-env:UserInstallation={(tmp_path / "profile").as_uri()}',
        '--headless', '--convert-to', 'docx:Office Open XML Text', '--outdir', str(out), str(local),
    ], capture_output=True, text=True, timeout=45)
    assert converted.returncode == 0, converted.stdout + converted.stderr
    result = (out / 'review.docx').read_bytes()
    document = xml(result, de.DOC)
    inserted = ''.join(n.text or '' for revision in document.iter(dc.q('w:ins')) for n in revision.iter(dc.q('w:t')))
    assert all(t in inserted for t in ('Replacement words', 'Inserted section', 'Inserted body'))
    assert ''.join(n.text or '' for n in document.iter(dc.q('w:delText'))) == 'alpha  beta gamma'
    assert len(list(document.iter(dc.q('w:drawing')))) == 2
    image_sizes = {Image.open(io.BytesIO(data)).size for name, data in parts(result).items() if name.startswith('word/media/') and name.endswith('.png')}
    assert {(100, 100), (100, 50)} <= image_sizes
