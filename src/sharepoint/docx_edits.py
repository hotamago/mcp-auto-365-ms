"""Conservative OOXML text revisions and raster image replacement, entirely offline."""

from __future__ import annotations

import copy
import hashlib
import io
import posixpath
import re
import warnings
import zipfile
from datetime import datetime
from pathlib import Path
from urllib.parse import unquote
from xml.etree import ElementTree as ET

from PIL import Image

from common.errors import Mcp365Error
from sharepoint import docx_comments as dc

W = dc.NS['w']
R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
PIC = 'http://schemas.openxmlformats.org/drawingml/2006/picture'
DOC = 'word/document.xml'
RELS = 'word/_rels/document.xml.rels'


def _target(base, target):
    target = unquote(target)
    return posixpath.normpath(target.lstrip('/') if target.startswith('/') else posixpath.join(base, target))


class Package:
    """Keep untouched ZIP payloads, member metadata, order and namespace bindings."""

    def __init__(self, data: bytes):
        try:
            with zipfile.ZipFile(io.BytesIO(data)) as archive:
                self.infos = archive.infolist()
                if len({i.filename for i in self.infos}) != len(self.infos):
                    dc.fail('DOCX contains duplicate ZIP entries.')
                self.original = {i.filename: archive.read(i) for i in self.infos}
                self.comment = archive.comment
            if any(n.startswith('_xmlsignatures/') for n in self.original):
                dc.fail('Signed DOCX files cannot be edited.')
            self.document = dc.XmlPart(self.original[DOC])
            types = dc.XmlPart(self.original['[Content_Types].xml']).root
            content_type = next((e.get('ContentType') for e in types if e.get('PartName') == '/word/document.xml'), '')
            if self.document.root.tag != dc.q('w:document') or content_type != 'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml':
                dc.fail('Only standard transitional .docx packages are supported.')
        except (zipfile.BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
            raise Mcp365Error(f'Invalid DOCX package: {exc}', dc.REMEDY) from exc

    def write(self, changed: dict[str, bytes]) -> bytes:
        remaining = dict(changed)
        out = io.BytesIO()
        with zipfile.ZipFile(out, 'w') as archive:
            archive.comment = self.comment
            for info in self.infos:
                archive.writestr(copy.copy(info), remaining.pop(info.filename, self.original[info.filename]))
            for name, data in remaining.items():
                archive.writestr(name, data, compress_type=zipfile.ZIP_DEFLATED)
        return out.getvalue()


def _plain(value, name, *, empty=False, tabs=False):
    if not isinstance(value, str) or (not empty and not value.strip()) or not dc._xml_safe(value):
        dc.fail(f'{name} must be XML-safe plain text' + ('.' if empty else ' and non-empty.'))
    if '\r' in value or ('\t' in value and not tabs):
        dc.fail(f'{name} cannot contain tabs or carriage returns.')
    return value


def _safe_paragraph(paragraph, parents):
    forbidden = {'ins', 'del', 'moveFrom', 'moveTo', 'moveFromRangeStart', 'moveFromRangeEnd', 'moveToRangeStart', 'moveToRangeEnd', 'fldSimple', 'fldChar', 'instrText', 'sdt', 'pPrChange', 'rPrChange', 'sectPr'}
    if any(e.tag in {dc.q('w:' + n) for n in forbidden} for e in paragraph.iter()):
        dc.fail('Target paragraph contains tracked revisions, complex fields, content controls or section properties.')
    node = paragraph
    while node is not None:
        if node.tag in {dc.q('w:' + n) for n in ('ins', 'del', 'moveFrom', 'moveTo', 'sdt', 'fldSimple')} or node.tag == dc.q('mc:AlternateContent'):
            dc.fail('Target paragraph is inside a revision, content control or alternate content.')
        node = parents.get(node)


def prepare_text(data: bytes, edits: list[dict], author: str) -> dc.EditResult:
    """Resolve every exact, case-sensitive match in the original before mutation."""
    _plain(author, 'author')
    if not isinstance(edits, list) or not edits:
        dc.fail('edits must be a non-empty list.')
    package = Package(data)
    document = package.document.root
    raw, locations, _, _, parents = dc._document_text(document)
    field_depth, field_text = 0, set()
    for node in document.iter():
        if node.tag == dc.q('w:fldChar'):
            kind = node.get(dc.q('w:fldCharType'))
            if kind == 'begin':
                field_depth += 1
            elif kind == 'end':
                field_depth = max(0, field_depth - 1)
        elif node.tag == dc.q('w:t') and field_depth:
            field_text.add(node)
    plans, previews = [], []
    for spec in edits:
        if not isinstance(spec, dict):
            dc.fail('Each edit must be an object.')
        insertion = 'insert_after' in spec
        allowed = {'insert_after', 'text', 'occurrence'} if insertion else {'find', 'replace', 'occurrence'}
        if set(spec) - allowed or not (allowed - {'occurrence'}) <= set(spec):
            dc.fail('Use {find, replace, occurrence?} or {insert_after, text, occurrence?}.')
        needle = _plain(spec['insert_after' if insertion else 'find'], 'anchor', tabs=True)
        replacement = _plain(spec['text' if insertion else 'replace'], 'replacement', empty=not insertion)
        if '\n' in needle or (not insertion and '\n' in replacement):
            dc.fail('Text replacement must stay in one paragraph; use insert_after for new paragraphs.')
        occurrence = spec.get('occurrence')
        if occurrence is not None and (type(occurrence) is not int or occurrence < 1):
            dc.fail('occurrence must be a positive 1-based integer.')
        matches = []
        for match in re.finditer(f'(?={re.escape(needle)})', raw):
            start, end = match.start(), match.start() + len(needle)
            selected = locations[start:end]
            if all(loc is not None and loc[2] is selected[0][2] for loc in selected):
                matches.append((start, end))
        if not matches:
            dc.fail(f'Text not found within one paragraph: {needle!r}.')
        if occurrence is None and len(matches) != 1:
            dc.fail(f'Ambiguous text: {len(matches)} matches of {needle!r}; specify occurrence.')
        if (occurrence or 1) > len(matches):
            dc.fail(f'Requested occurrence {occurrence}, found only {len(matches)}.')
        start, end = matches[(occurrence or 1) - 1]
        paragraph = locations[start][2]
        if any(loc[0] in field_text for loc in locations[start:end]):
            dc.fail('Target intersects a complex field result, possibly spanning paragraphs.')
        _safe_paragraph(paragraph, parents)
        selected_runs = {parents[loc[0]] for loc in locations[start:end]}
        if any(r.tag != dc.q('w:r') or parents.get(r) is not paragraph for r in selected_runs):
            dc.fail('Target must contain ordinary runs directly in one paragraph; hyperlinks are unsupported.')
        if any(any(c.tag not in {dc.q('w:rPr'), dc.q('w:t'), dc.q('w:tab')} for c in r) for r in selected_runs):
            dc.fail('Target intersects a complex run.')
        # Reject fields/objects between endpoints as well as those carrying text.
        first, last = parents[locations[start][0]], parents[locations[end - 1][0]]
        children = list(paragraph)
        markers = {dc.q('w:' + n) for n in ('bookmarkStart', 'bookmarkEnd', 'commentRangeStart', 'commentRangeEnd', 'proofErr')}
        for node in children[children.index(first):children.index(last) + 1]:
            if node.tag != dc.q('w:r') and node.tag not in markers:
                dc.fail('Target range contains unsupported markup.')
            if node.tag == dc.q('w:r') and any(c.tag not in {dc.q('w:rPr'), dc.q('w:t'), dc.q('w:tab'), dc.q('w:commentReference')} for c in node):
                dc.fail('Target range crosses a field, drawing or complex run.')
        for old in plans:
            if old['paragraph'] is paragraph and (insertion or old['insertion'] or max(start, old['start']) < min(end, old['end'])):
                dc.fail('Overlapping edits or multiple operations on an insertion anchor paragraph are unsupported.')
        props = first.find('w:rPr', dc.NS)
        plans.append(dict(start=start, end=end, paragraph=paragraph, insertion=insertion, text=replacement, props=props))
        previews.append({**spec, 'occurrence': occurrence or 1, 'paragraph_text': ''.join(t.text or '' for t in paragraph.iter(dc.q('w:t')))})
    ids = [int(e.get(dc.q('w:id'))) for e in document.iter() if (e.get(dc.q('w:id')) or '').lstrip('-').isdigit()]
    next_id = max(ids, default=-1) + 1
    date = datetime.now().isoformat(timespec='seconds')
    records = []

    def revision(kind):
        nonlocal next_id
        node = ET.Element(dc.q('w:' + kind), {dc.q('w:id'): str(next_id), dc.q('w:author'): author, dc.q('w:date'): date})
        next_id += 1
        return node

    def text_run(text, props):
        run = ET.Element(dc.q('w:r'))
        if props is not None:
            run.append(copy.deepcopy(props))
        ET.SubElement(run, dc.q('w:t'), {dc.q('xml:space'): 'preserve'}).text = text
        return run

    # Right-to-left keeps coordinates in the original valid as deleted text leaves
    # the visible text stream. New paragraphs are inserted only after replacements.
    for plan in sorted((p for p in plans if not p['insertion']), key=lambda p: p['start'], reverse=True):
        container, at = dc._boundary(document, plan['end'] - 1, True)
        sentinel = ET.Element('end-of-edit')
        container.insert(at, sentinel)
        paragraph, begin = dc._boundary(document, plan['start'], False)
        finish = list(paragraph).index(sentinel)
        nodes = list(paragraph)[begin:finish]
        paragraph.remove(sentinel)
        inserted = revision('ins') if plan['text'] else None
        if inserted is not None:
            inserted.append(text_run(plan['text'], plan['props']))
        changed = []
        for node in nodes:
            paragraph.remove(node)
            if node.tag == dc.q('w:r') and any(n.tag in {dc.q('w:t'), dc.q('w:tab')} for n in node):
                deleted = revision('del')
                for t in node.iter(dc.q('w:t')):
                    t.tag = dc.q('w:delText')
                deleted.append(node)
                changed.append(deleted)
            else:
                # Existing comments/bookmarks/reference runs keep their identity.
                changed.append(node)
        if inserted is not None:
            changed.append(inserted)
        for offset, node in enumerate(changed):
            paragraph.insert(begin + offset, node)
    for plan in (p for p in plans if p['insertion']):
        anchor = plan['paragraph']
        container = parents[anchor]
        at = list(container).index(anchor) + 1
        for line in plan['text'].split('\n'):
            paragraph = ET.Element(dc.q('w:p'))
            original_props = anchor.find('w:pPr', dc.NS)
            pprops = copy.deepcopy(original_props) if original_props is not None else ET.Element(dc.q('w:pPr'))
            paragraph.append(pprops)
            mark_props = pprops.find('w:rPr', dc.NS)
            if mark_props is None:
                mark_props = ET.SubElement(pprops, dc.q('w:rPr'))
            # Track the paragraph mark too: rejecting removes the whole paragraph.
            mark_props.append(revision('ins'))
            inserted = revision('ins')
            inserted.append(text_run(line, plan['props']))
            paragraph.append(inserted)
            container.insert(at, paragraph)
            at += 1
    original_ids = {e.get(dc.q('w:id')) for e in ET.fromstring(package.original[DOC]).iter()}
    for node in document.iter():
        if node.tag in {dc.q('w:ins'), dc.q('w:del')} and node.get(dc.q('w:id')) not in original_ids:
            records.append({'xml': ET.tostring(node, encoding='unicode'), 'id': node.get(dc.q('w:id'))})
    result = dc.EditResult(package.write({DOC: package.document.bytes()}), previews, records)
    verify_text(result.data, records)
    return result


def verify_text(data, records):
    document = Package(data).document.root
    for record in records:
        found = [n for n in document.iter() if n.tag in {dc.q('w:ins'), dc.q('w:del')} and n.get(dc.q('w:id')) == record['id']]
        if len(found) != 1 or ET.tostring(found[0], encoding='unicode') != record['xml']:
            dc.fail(f'Tracked revision {record["id"]} failed verification.')


def image_info(data: bytes) -> dict:
    """Decode the entire PNG/JPEG, including CRC checks, before packaging it."""
    try:
        with warnings.catch_warnings():
            warnings.simplefilter('error', Image.DecompressionBombWarning)
            with Image.open(io.BytesIO(data)) as image:
                if image.format not in {'PNG', 'JPEG'} or getattr(image, 'n_frames', 1) != 1:
                    dc.fail('Only single-frame PNG/JPEG images are supported.')
                fmt, size = image.format, image.size
                image.verify()
            with Image.open(io.BytesIO(data)) as image:
                image.load()
                if image.getexif().get(274, 1) != 1:
                    dc.fail('EXIF-rotated images must be normalized locally first.')
        return {'format': fmt, 'width': size[0], 'height': size[1], 'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}
    except (OSError, ValueError, SyntaxError, Image.DecompressionBombError, Image.DecompressionBombWarning) as exc:
        raise Mcp365Error(f'Invalid PNG/JPEG image: {exc}', dc.REMEDY) from exc


def locate_image(package: Package, locate: dict):
    if not isinstance(locate, dict) or len(locate) != 1 or not set(locate) <= {'caption_text', 'image_name'}:
        dc.fail('locate requires exactly one of caption_text or image_name.')
    value = _plain(next(iter(locate.values())), 'image locator')
    root = package.document.root
    parents = {c: p for p in root.iter() for c in p}
    rels = dc.XmlPart(package.original[RELS])
    by_id = {e.get('Id'): e for e in rels.root}
    if len(by_id) != len(rels.root):
        dc.fail('Duplicate document relationship IDs.')
    candidates = []
    for drawing in root.iter(dc.q('w:drawing')):
        blips = list(drawing.iter(f'{{{A}}}blip'))
        if len(blips) != 1:
            continue
        rid = blips[0].get(f'{{{R}}}embed')
        rel = by_id.get(rid)
        if rel is None or rel.get('TargetMode') == 'External' or not rel.get('Type', '').endswith('/image'):
            continue
        target = rel.get('Target', '')
        media = _target('word', target)
        if not media.startswith(('word/media/', 'media/')) or media not in package.original:
            dc.fail('Image relationship has an unsupported media target.')
        names = {media, posixpath.basename(media)}
        for node in drawing.iter():
            if node.tag in {f'{{{WP}}}docPr', f'{{{PIC}}}cNvPr'}:
                names.update(node.get(k, '') for k in ('name', 'descr', 'title'))
        matched = value in names if 'image_name' in locate else False
        if 'caption_text' in locate:
            paragraph = parents.get(drawing)
            while paragraph is not None and paragraph.tag != dc.q('w:p'):
                paragraph = parents.get(paragraph)
            if paragraph is not None and paragraph in parents:
                siblings = list(parents[paragraph])
                at = siblings.index(paragraph)
                adjacent = siblings[max(0, at - 1):at] + siblings[at + 1:at + 2]
                matched = any(p.tag == dc.q('w:p') and value in ''.join(t.text or '' for t in p.iter(dc.q('w:t'))) for p in adjacent)
        if matched:
            # Updating drawings inside tracked changes/AlternateContent can alter
            # both current and historical or fallback representations.
            node = drawing
            while node is not None:
                if node.tag in {dc.q('w:' + n) for n in ('ins', 'del', 'moveFrom', 'moveTo')} or node.tag == dc.q('mc:AlternateContent'):
                    dc.fail('Target image is inside a revision or alternate content.')
                node = parents.get(node)
            candidates.append((drawing, blips[0], rel, media, rels))
    if len(candidates) != 1:
        dc.fail(f'Image locator matched {len(candidates)} drawings; use a unique caption or image name.')
    return candidates[0]


def prepare_image(data: bytes, locate: dict, new_image: bytes) -> tuple[dc.EditResult, bytes]:
    package = Package(data)
    drawing, blip, rel, media, rels = locate_image(package, locate)
    current = package.original[media]
    old_info, new_info = image_info(current), image_info(new_image)
    extents = list(drawing.iter(f'{{{WP}}}extent'))
    if len(extents) != 1:
        dc.fail('Target drawing has no unique displayed extent.')
    extent = extents[0]
    cx, cy = int(extent.get('cx', 0)), int(extent.get('cy', 0))
    if cx <= 0 or cy <= 0:
        dc.fail('Target drawing has invalid displayed extent.')
    # Compare intrinsic old/new ratios; preserve deliberately stretched drawings
    # unless the replacement changes aspect ratio by more than two percent.
    ratio_changed = abs((new_info['width'] / new_info['height']) / (old_info['width'] / old_info['height']) - 1) > .02
    new_cy = round(cx * new_info['height'] / new_info['width']) if ratio_changed else cy
    if ratio_changed:
        extent.set('cy', str(new_cy))
        for transform in drawing.iter(f'{{{A}}}xfrm'):
            for shape_extent in transform.findall(f'{{{A}}}ext'):
                shape_extent.set('cx', str(cx))
                shape_extent.set('cy', str(new_cy))
    # Count actual uses of this relationship and aliases from every XML part,
    # including headers/footers/VML and another rId for the same media part.
    shared = False
    for name, content in package.original.items():
        if not name.endswith('.rels'):
            continue
        relationships = dc.XmlPart(content).root
        base = posixpath.dirname(posixpath.dirname(name))
        for r in relationships:
            target = r.get('Target', '')
            resolved = _target(base, target)
            if r.get('TargetMode') != 'External' and resolved == media:
                if name != RELS or r.get('Id') != rel.get('Id'):
                    shared = True
    rid = rel.get('Id')
    uses = sum(n.get(f'{{{R}}}embed') == rid or n.get(f'{{{R}}}id') == rid or n.get(f'{{{R}}}link') == rid for n in package.document.root.iter())
    shared |= uses != 1
    changed = {DOC: package.document.bytes()}
    different_format = old_info['format'] != new_info['format']
    output_media = media
    if shared or different_format:
        suffix = 'png' if new_info['format'] == 'PNG' else 'jpg'
        index = 1
        while f'word/media/mcp-image-{index}.{suffix}' in package.original:
            index += 1
        output_media = f'word/media/mcp-image-{index}.{suffix}'
        if uses > 1:
            # A shared rId cannot stay unchanged without replacing other drawings.
            # Clone it for this drawing only; all existing relationships survive.
            used = {r.get('Id') for r in rels.root}
            index = 1
            while f'rId{index}' in used:
                index += 1
            rel = copy.deepcopy(rel)
            rel.set('Id', f'rId{index}')
            rels.root.append(rel)
            blip.set(f'{{{R}}}embed', rel.get('Id'))
        rel.set('Target', posixpath.relpath(output_media, 'word'))
        changed[RELS] = rels.bytes()
        types = dc.XmlPart(package.original['[Content_Types].xml'])
        # An override avoids relying on any possibly conflicting extension default.
        ET.SubElement(types.root, dc.q('ct:Override'), {'PartName': '/' + output_media, 'ContentType': 'image/png' if suffix == 'png' else 'image/jpeg'})
        changed['[Content_Types].xml'] = types.bytes()
    changed[DOC] = package.document.bytes()
    changed[output_media] = new_image
    preview = [{
        'locate': locate, 'current_media': media, 'current_image': old_info,
        'replacement_image': new_info, 'shared_media': shared,
        'relationship_id_before': rid, 'relationship_id_after': rel.get('Id'),
        'output_media': output_media, 'extent_before': {'cx': cx, 'cy': cy},
        'extent_after': {'cx': cx, 'cy': new_cy},
        'aspect_ratio_changed_over_2_percent': ratio_changed,
        'review_note': 'Image replacement is not a Word tracked revision; inspect the old/new images before approval.',
    }]
    result = dc.EditResult(package.write(changed), preview, [])
    # Verify all changed payloads after re-download, including the drawing,
    # relationships, content types and raster bytes, not just their presence.
    result.records = [{'part': n, 'sha256': hashlib.sha256(v).hexdigest()} for n, v in changed.items()]
    verify_image(result.data, result.records)
    return result, current


def verify_image(data, records):
    package = Package(data)
    for record in records:
        if hashlib.sha256(package.original.get(record['part'], b'')).hexdigest() != record['sha256']:
            dc.fail(f'Image replacement part {record["part"]} failed verification.')


def read_image(path: str) -> bytes:
    source = Path(path).expanduser()
    if source.suffix.lower() not in {'.png', '.jpg', '.jpeg'} or not source.is_file():
        dc.fail('new_image_path must name an existing local PNG/JPEG file.')
    data = source.read_bytes()
    info = image_info(data)
    if (source.suffix.lower() == '.png') != (info['format'] == 'PNG'):
        dc.fail('Image filename extension does not match the decoded format.')
    return data
