"""Lossless package editing for Word review comments. No network or document model.

Only document/comment XML and required package declarations change. Unmodified
ZIP member payloads, ordering, compression methods and metadata are retained.
Unsupported anchors fail closed rather than flattening fields or revisions.
"""

from __future__ import annotations

import copy
import io
import re
import secrets
import zipfile
from dataclasses import dataclass
from datetime import UTC, datetime
from html import escape
from typing import Any
from xml.etree import ElementTree as ET

from common.errors import Mcp365Error

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "w14": "http://schemas.microsoft.com/office/word/2010/wordml",
    "w15": "http://schemas.microsoft.com/office/word/2012/wordml",
    "w16cid": "http://schemas.microsoft.com/office/word/2016/wordml/cid",
    "w16cex": "http://schemas.microsoft.com/office/word/2018/wordml/cex",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "rel": "http://schemas.openxmlformats.org/package/2006/relationships",
    "xml": "http://www.w3.org/XML/1998/namespace",
}
PARTS = {
    "comments": ("w", "comments", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"),
    "commentsExtended": ("w15", "commentsEx", "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"),
    "commentsIds": ("w16cid", "commentsIds", "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds"),
    "commentsExtensible": ("w16cex", "commentsExtensible", "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible"),
}
REMEDY = "Nothing uploaded. Check the anchor/thread and re-run the dry run; use Word Online for unsupported structures."


def q(name: str) -> str:
    prefix, local = name.split(":")
    return f"{{{NS[prefix]}}}{local}"


def fail(message: str) -> None:
    raise Mcp365Error(message, REMEDY)


def _attribute(value: str) -> str:
    # Literal whitespace in attributes is normalized on parse; character
    # references preserve existing values such as multiline drawing metadata.
    return escape(value, quote=True).replace("\t", "&#9;").replace("\n", "&#10;").replace("\r", "&#13;")


def _xml_safe(value: str) -> bool:
    return all(c in "\t\r\n" or 0x20 <= ord(c) <= 0xD7FF or 0xE000 <= ord(c) <= 0xFFFD or 0x10000 <= ord(c) <= 0x10FFFF for c in value)


class XmlPart:
    """ElementTree with all original prefix declarations, including mc:Ignorable.

    ET's serializer drops unused declarations even when prefix names occur in
    attribute values. A local serializer retains them without global namespace
    registration (tools may edit different files concurrently).
    """

    def __init__(self, data: bytes):
        if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
            fail("DTD/entity declarations in DOCX XML are unsupported.")
        # Keep declarations on their actual elements so scoped rebinding and
        # QName-valued mc attributes survive. These plain xmlns attributes also
        # survive deepcopy when an anchored run is split.
        pending = {}
        parser = ET.XMLParser(target=ET.TreeBuilder(insert_comments=True, insert_pis=True))
        parsed = ET.iterparse(io.BytesIO(data), events=("start-ns", "start"), parser=parser)
        for event, value in parsed:
            if event == "start-ns":
                prefix, uri = value
                pending["xmlns:" + prefix if prefix else "xmlns"] = uri
            else:
                value.attrib.update(pending)
                pending = {}
        self.root = parsed.root

    @classmethod
    def new(cls, name: str) -> XmlPart:
        prefix, local = name.split(":")
        return cls(f'<{prefix}:{local} xmlns:{prefix}="{NS[prefix]}"/>'.encode())

    def bytes(self) -> bytes:
        def render(node, inherited):
            if node.tag is ET.Comment:
                return f"<!--{node.text or ''}-->" + escape(node.tail or "", quote=False)
            if node.tag is ET.ProcessingInstruction:
                return f"<?{node.text or ''}?>" + escape(node.tail or "", quote=False)
            declared = {
                "" if k == "xmlns" else k[6:]: v
                for k, v in node.attrib.items() if k == "xmlns" or k.startswith("xmlns:")
            }
            bindings = {**inherited, **declared}

            def tag(name, attribute=False):
                if not name.startswith("{"):
                    # A no-namespace element inside a default namespace needs
                    # an explicit reset to keep its expanded name unchanged.
                    if not attribute and bindings.get(""):
                        bindings[""] = declared[""] = ""
                    return name
                uri, local = name[1:].split("}")
                matches = [p for p, u in bindings.items() if u == uri and (p or not attribute)]
                if matches:
                    prefix = matches[0]
                else:
                    prefix = next((p for p, u in NS.items() if u == uri), f"ns{len(bindings)}")
                    while prefix in bindings and bindings[prefix] != uri:
                        prefix += "x"
                    bindings[prefix] = declared[prefix] = uri
                return f"{prefix}:{local}" if prefix else local

            element_name = tag(node.tag)
            attrs = [(tag(k, True), v) for k, v in node.attrib.items() if k != "xmlns" and not k.startswith("xmlns:")]
            declarations = "".join(f' xmlns{":" + p if p else ""}="{_attribute(u)}"' for p, u in declared.items())
            attributes = "".join(f' {k}="{_attribute(v)}"' for k, v in attrs)
            inner = escape(node.text or "", quote=False) + "".join(render(child, bindings) for child in node)
            return f"<{element_name}{declarations}{attributes}>{inner}</{element_name}>" + escape(node.tail or "", quote=False)

        return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + render(self.root, {"xml": NS["xml"]})).encode()


@dataclass
class EditResult:
    data: bytes
    preview: list[dict[str, Any]]
    records: list[dict[str, str]]


def _normalize(value: str) -> str:
    return " ".join(value.split())


def _text(comment) -> str:
    return "\n".join("".join(t.text or "" for t in p.iter(q("w:t"))) for p in comment.findall("w:p", NS))


def _document_text(document):
    """Visible text + character locations; paragraph separators normalize to spaces."""
    parents = {c: p for p in document.iter() for c in p}
    chars, locations = [], []
    for paragraph in document.iter(q("w:p")):
        for node in paragraph.iter():
            if node.tag not in (q("w:t"), q("w:tab"), q("w:br"), q("w:cr")):
                continue
            # Exclude nested textbox paragraphs from their containing paragraph.
            ancestor = parents.get(node)
            while ancestor is not None and ancestor.tag != q("w:p"):
                ancestor = parents.get(ancestor)
            if ancestor is not paragraph:
                continue
            value = node.text or "" if node.tag == q("w:t") else "\t" if node.tag == q("w:tab") else "\n"
            for offset, char in enumerate(value):
                chars.append(char)
                locations.append((node, offset, paragraph))
        chars.append("\n")
        locations.append(None)
    raw = "".join(chars)
    normalized, spans = [], []
    for match in re.finditer(r"\S|\s+", raw):
        normalized.append(" " if match[0].isspace() else match[0])
        spans.append((match.start(), match.end()))
    return raw, locations, "".join(normalized), spans, parents


def _resolve_anchor(document, spec):
    _, locations, normalized, spans, parents = _document_text(document)
    needle = _normalize(spec["anchor_text"])
    matches = list(re.finditer(f"(?={re.escape(needle)})", normalized))
    occurrence = spec.get("occurrence")
    if occurrence is not None and (type(occurrence) is not int or occurrence < 1):
        fail("occurrence must be a positive 1-based integer.")
    if not matches:
        fail(f"Anchor not found: {spec['anchor_text']!r}.")
    if occurrence is None and len(matches) != 1:
        fail(f"Ambiguous anchor: {len(matches)} occurrences of {spec['anchor_text']!r}; specify occurrence.")
    if (occurrence or 1) > len(matches):
        fail(f"Requested occurrence {occurrence}, found only {len(matches)}.")
    match = matches[(occurrence or 1) - 1]
    start, end = spans[match.start()][0], spans[match.start() + len(needle) - 1][1]
    forbidden = {q(f"w:{s}") for s in ("ins", "del", "moveFrom", "moveTo", "fldSimple", "sdt", "drawing", "pict")}
    field_depth, field_text = 0, set()
    for node in document.iter():
        if node.tag == q("w:fldChar"):
            kind = node.get(q("w:fldCharType"))
            if kind == "begin":
                field_depth += 1
            elif kind == "end":
                field_depth = max(0, field_depth - 1)
        elif node.tag == q("w:t") and field_depth:
            field_text.add(node)
    runs = set()
    for location in locations[start:end]:
        if location is None:
            continue
        node = location[0]
        if node in field_text:
            fail("Anchor intersects a complex field result.")
        run = parents.get(node)
        if run is None or run.tag != q("w:r"):
            fail("Anchor text is not in ordinary Word runs.")
        ancestor = run
        while ancestor is not None:
            if ancestor.tag in forbidden or ancestor.tag == f"{{{NS['mc']}}}AlternateContent":
                fail("Anchor intersects a tracked revision, field, content control or alternate content.")
            ancestor = parents.get(ancestor)
        runs.add(run)
    allowed = {q(f"w:{s}") for s in ("rPr", "t", "tab", "br", "cr")}
    if any(any(c.tag not in allowed for c in run) for run in runs):
        fail("Anchor intersects a complex run (field, drawing or reference).")
    first, last = locations[start], locations[end - 1]
    if first is None or last is None:
        fail("Anchor must begin and end in text.")
    excerpt = _normalize("".join(first[2].itertext()))
    return start, end, excerpt[:300]


def _boundary(document, position: int, after: bool):
    _, locations, _, _, parents = _document_text(document)
    text_node, offset, _ = locations[position]
    offset += int(after)
    run = parents[text_node]
    container = parents[run]
    children = list(run)
    index = children.index(text_node)
    value = text_node.text or "" if text_node.tag == q("w:t") else "x"
    left, right = copy.deepcopy(run), copy.deepcopy(run)
    left[:] = [copy.deepcopy(c) for c in children[:index] if c.tag != q("w:rPr")]
    right[:] = [copy.deepcopy(c) for c in children[index + 1:] if c.tag != q("w:rPr")]
    for side in (left, right):
        if run.find("w:rPr", NS) is not None:
            side.insert(0, copy.deepcopy(run.find("w:rPr", NS)))
        side.tail = None
    if offset:
        node = copy.deepcopy(text_node)
        if node.tag == q("w:t"):
            node.text = value[:offset]
            node.set(q("xml:space"), "preserve")
        left.append(node)
    if offset < len(value):
        node = copy.deepcopy(text_node)
        if node.tag == q("w:t"):
            node.text = value[offset:]
            node.set(q("xml:space"), "preserve")
        right.insert(1 if right.find("w:rPr", NS) is not None else 0, node)
    at = list(container).index(run)
    container.remove(run)
    left_has_content = any(c.tag != q("w:rPr") for c in left)
    right_has_content = any(c.tag != q("w:rPr") for c in right)
    if left_has_content:
        container.insert(at, left)
        at += 1
    if right_has_content:
        container.insert(at, right)
        right.tail = run.tail
    elif left_has_content:
        left.tail = run.tail
    return container, at


def _reference(comment_id):
    run = ET.Element(q("w:r"))
    props = ET.SubElement(run, q("w:rPr"))
    ET.SubElement(props, q("w:rStyle"), {q("w:val"): "CommentReference"})
    ET.SubElement(run, q("w:commentReference"), {q("w:id"): comment_id})
    return run


def _mark(document, start, end, comment_id):
    container, at = _boundary(document, end - 1, True)
    container.insert(at, ET.Element(q("w:commentRangeEnd"), {q("w:id"): comment_id}))
    container.insert(at + 1, _reference(comment_id))
    container, at = _boundary(document, start, False)
    container.insert(at, ET.Element(q("w:commentRangeStart"), {q("w:id"): comment_id}))


def _mirror(document, parent_id, comment_id):
    parents = {c: p for p in document.iter() for c in p}
    found = {}
    for name in ("commentRangeStart", "commentRangeEnd", "commentReference"):
        nodes = [e for e in document.iter(q(f"w:{name}")) if e.get(q("w:id")) == parent_id]
        if len(nodes) != 1:
            fail(f"Reply target {parent_id} has no unique document anchor.")
        found[name] = nodes[0]
    for name in ("commentRangeStart", "commentRangeEnd"):
        original = found[name]
        container = parents[original]
        container.insert(list(container).index(original) + 1, ET.Element(q(f"w:{name}"), {q("w:id"): comment_id}))
    run = parents[found["commentReference"]]
    container = parents[run]
    container.insert(list(container).index(run) + 1, _reference(comment_id))


def _hex_id(taken):
    while True:
        value = f"{secrets.randbelow(0x7FFFFFFE) + 1:08X}"
        if value not in taken:
            taken.add(value)
            return value


def _prepare(data: bytes, comments: list[dict[str, Any]], author: str) -> EditResult:
    if not author.strip() or not _xml_safe(author):
        fail("A non-empty comment author is required.")
    if not isinstance(comments, list) or not comments:
        fail("comments must be a non-empty list.")
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        infos = archive.infolist()
        if len({i.filename for i in infos}) != len(infos):
            fail("DOCX contains duplicate ZIP entries.")
        original = {i.filename: archive.read(i) for i in infos}
        archive_comment = archive.comment
    if any(n.startswith("_xmlsignatures/") for n in original):
        fail("Signed DOCX files cannot be changed without invalidating their signature.")
    xml = {n: XmlPart(original[n]) for n in ("word/document.xml", "[Content_Types].xml")}
    rel_name = "word/_rels/document.xml.rels"
    # OPC consumers (including LibreOffice) expect package relationships to
    # use the default namespace, although prefixed XML is namespace-equivalent.
    xml[rel_name] = XmlPart(original[rel_name]) if rel_name in original else XmlPart(
        f'<Relationships xmlns="{NS["rel"]}"/>'.encode()
    )
    for name, (prefix, root, _) in PARTS.items():
        path = f"word/{name}.xml"
        if path in original:
            xml[path] = XmlPart(original[path])
        elif name in ("comments", "commentsExtended"):
            xml[path] = XmlPart.new(f"{prefix}:{root}")
        if path in xml and xml[path].root.tag != q(f"{prefix}:{root}"):
            fail(f"Unexpected XML root in {path}.")
    if xml[rel_name].root.tag != q("rel:Relationships") or xml["[Content_Types].xml"].root.tag != q("ct:Types"):
        fail("Invalid DOCX package declarations.")
    document = xml["word/document.xml"].root
    if document.tag != q("w:document"):
        fail("Only transitional Word DOCX documents are supported.")
    types = xml["[Content_Types].xml"].root
    doc_type = next((e.get("ContentType") for e in types if e.get("PartName") == "/word/document.xml"), "")
    if doc_type != "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml":
        fail("The package is not a standard .docx document.")
    existing = xml["word/comments.xml"].root
    # Make the Word 2010 paragraph identifiers ignorable to older consumers.
    prefix = "w14"
    while existing.get(f"xmlns:{prefix}", NS["w14"]) != NS["w14"]:
        prefix += "x"
    existing.set(f"xmlns:{prefix}", NS["w14"])
    ignorable = existing.get(q("mc:Ignorable"), "").split()
    if prefix not in ignorable:
        existing.set(q("mc:Ignorable"), " ".join([*ignorable, prefix]))
    existing_comments = existing.findall("w:comment", NS)
    by_id = {c.get(q("w:id")): c for c in existing_comments}
    if len(by_id) != len(existing_comments) or None in by_id:
        fail("Existing comment IDs are not unique.")
    taken = {v.upper() for part in xml.values() for node in part.root.iter() for k, v in node.attrib.items() if k in (q("w14:paraId"), q("w15:paraId"), q("w16cid:paraId"), q("w16cid:durableId"), q("w16cex:durableId"))}
    ids = {int(v) for v in by_id}
    ids.update(int(e.get(q("w:id"))) for e in document.iter() if e.tag in (q("w:commentRangeStart"), q("w:commentRangeEnd"), q("w:commentReference")))
    next_id = max(ids, default=-1) + 1
    plans, previews = [], []
    legacy_paragraph_ids = {}
    # Resolve the entire batch against the original document before any mutation.
    for spec in comments:
        if not isinstance(spec, dict) or set(spec) - {"anchor_text", "text", "occurrence", "reply_to"}:
            fail("Each comment must contain only anchor_text, text, occurrence and/or reply_to.")
        text = spec.get("text")
        if not isinstance(text, str) or not text.strip() or not _xml_safe(text):
            fail("Comment text must be non-empty XML-safe plain text.")
        anchor, reply = spec.get("anchor_text"), spec.get("reply_to")
        if bool(anchor) == (reply is not None):
            fail("Specify exactly one of anchor_text or reply_to.")
        if reply is not None:
            if "occurrence" in spec:
                fail("occurrence applies only to text anchors.")
            target = str(reply)
            candidates = [by_id[target]] if target in by_id else [c for c in existing_comments if _normalize(target).casefold() in _normalize(f"{c.get(q('w:author'), '')}: {_text(c)}").casefold()]
            if len(candidates) != 1 or not target.strip():
                fail(f"Reply target {target!r} not found or ambiguous; use an existing w:id or unique 'Author: text' substring.")
            parent = candidates[0]
            paragraphs = parent.findall("w:p", NS)
            if not paragraphs:
                fail("Reply target has no paragraph.")
            last = paragraphs[-1]
            parent_para_id = last.get(q("w14:paraId"))
            if not parent_para_id:
                # Add only the threading identifier to legacy comments; leave
                # all of their text, author, date and anchor unchanged.
                if last not in legacy_paragraph_ids:
                    legacy_paragraph_ids[last] = _hex_id(taken)
                parent_para_id = legacy_paragraph_ids[last]
            parent_id = parent.get(q("w:id"))
            plans.append(("reply", parent_id, parent_para_id, text))
            previews.append({"reply_to": parent_id, "parent_author": parent.get(q("w:author"), ""), "parent_text": _text(parent), "text": text})
        else:
            if not isinstance(anchor, str) or not _normalize(anchor):
                fail("anchor_text must be non-empty text.")
            start, end, excerpt = _resolve_anchor(document, spec)
            plans.append(("anchor", start, end, text))
            previews.append({"anchor_text": anchor, "occurrence": spec.get("occurrence", 1), "paragraph_excerpt": excerpt, "text": text})
    now = datetime.now(UTC)
    # Word writes w:date as naive local time (verified on Word Online comments); dateUtc stays UTC.
    date = now.astimezone().replace(tzinfo=None).isoformat(timespec="seconds")
    date_utc = now.isoformat(timespec="milliseconds").replace("+00:00", "Z")
    initials = "".join(word[0] for word in author.split("(")[0].split()).upper()
    records = []
    for paragraph, para_id in legacy_paragraph_ids.items():
        paragraph.set(q("w14:paraId"), para_id)
    extended = xml["word/commentsExtended.xml"].root
    # Populate commentEx for legacy top-level comments that already have paraIds.
    ex_ids = {e.get(q("w15:paraId")) for e in extended}
    for old in existing_comments:
        paragraphs = old.findall("w:p", NS)
        para = paragraphs[-1].get(q("w14:paraId")) if paragraphs else None
        if para and para not in ex_ids:
            ET.SubElement(extended, q("w15:commentEx"), {q("w15:paraId"): para, q("w15:done"): "0"})
    for index, (kind, first, second, text) in enumerate(plans):
        comment_id = str(next_id + index)
        node = ET.SubElement(existing, q("w:comment"), {q("w:id"): comment_id, q("w:author"): author, q("w:initials"): initials, q("w:date"): date})
        para_id = ""
        for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
            para_id = _hex_id(taken)
            para = ET.SubElement(node, q("w:p"), {q("w14:paraId"): para_id})
            run = ET.SubElement(para, q("w:r"))
            ET.SubElement(run, q("w:t"), {q("xml:space"): "preserve"}).text = line
        attrs = {q("w15:paraId"): para_id, q("w15:done"): "0"}
        if kind == "reply":
            attrs[q("w15:paraIdParent")] = second
            _mirror(document, first, comment_id)
        else:
            _mark(document, first, second, comment_id)
        ET.SubElement(extended, q("w15:commentEx"), attrs)
        durable = _hex_id(taken)
        if "word/commentsExtensible.xml" in xml and "word/commentsIds.xml" not in xml:
            fail("commentsExtensible exists without commentsIds; repair in Word first.")
        if "word/commentsIds.xml" in xml:
            ET.SubElement(xml["word/commentsIds.xml"].root, q("w16cid:commentId"), {q("w16cid:paraId"): para_id, q("w16cid:durableId"): durable})
        if "word/commentsExtensible.xml" in xml:
            ET.SubElement(xml["word/commentsExtensible.xml"].root, q("w16cex:commentExtensible"), {q("w16cex:durableId"): durable, q("w16cex:dateUtc"): date_utc})
        records.append({
            "id": comment_id, "text": _text(node), "author": author,
            "para_id": para_id, "parent_para_id": second if kind == "reply" else "",
            "parent_id": first if kind == "reply" else "",
            "anchor_text": previews[index].get("anchor_text", ""),
            "durable_id": durable if "word/commentsIds.xml" in xml else "",
        })
        previews[index]["new_comment_id"] = comment_id
    rels = xml[rel_name].root
    for name, (_, _, relationship) in PARTS.items():
        path = f"word/{name}.xml"
        if path not in xml:
            continue
        matching = [r for r in rels if r.get("Type") == relationship]
        if matching and (len(matching) != 1 or matching[0].get("Target") not in (f"{name}.xml", f"/word/{name}.xml") or matching[0].get("TargetMode") == "External"):
            fail(f"Nonstandard {name} relationship target is unsupported.")
        if not matching:
            used = {r.get("Id") for r in rels}
            rid = 1
            while f"rId{rid}" in used:
                rid += 1
            ET.SubElement(rels, q("rel:Relationship"), {"Id": f"rId{rid}", "Type": relationship, "Target": f"{name}.xml"})
        content_type = f"application/vnd.openxmlformats-officedocument.wordprocessingml.{name}+xml"
        declarations = [t for t in types if t.get("PartName") == f"/{path}"]
        if declarations and (len(declarations) != 1 or declarations[0].get("ContentType") != content_type):
            fail(f"Unexpected content type for {name}.")
        if not declarations:
            ET.SubElement(types, q("ct:Override"), {"PartName": f"/{path}", "ContentType": content_type})
    changed = {name: part.bytes() for name, part in xml.items()}
    # Avoid rewriting declarations that already contain everything required.
    for name in ("[Content_Types].xml", rel_name):
        if name in original and ET.tostring(XmlPart(original[name]).root) == ET.tostring(xml[name].root):
            changed.pop(name)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as archive:
        archive.comment = archive_comment
        for info in infos:
            archive.writestr(copy.copy(info), changed.pop(info.filename, original[info.filename]))
        for name, content in changed.items():
            archive.writestr(name, content, compress_type=zipfile.ZIP_DEFLATED)
    result = EditResult(output.getvalue(), previews, records)
    verify(result.data, records)
    return result


def prepare(data: bytes, comments: list[dict[str, Any]], author: str) -> EditResult:
    """Validate the complete batch, then produce a new in-memory DOCX package."""
    try:
        return _prepare(data, comments, author)
    except (zipfile.BadZipFile, ET.ParseError, KeyError, ValueError) as exc:
        raise Mcp365Error(f"Invalid or unsupported DOCX package: {exc}", REMEDY) from exc


def verify(data: bytes, records: list[dict[str, str]]) -> None:
    """Verify comment content, identity, thread linkage and document references."""
    with zipfile.ZipFile(io.BytesIO(data)) as archive:
        comments = ET.fromstring(archive.read("word/comments.xml"))
        document = ET.fromstring(archive.read("word/document.xml"))
        extended = ET.fromstring(archive.read("word/commentsExtended.xml"))
        identifiers = ET.fromstring(archive.read("word/commentsIds.xml")) if "word/commentsIds.xml" in archive.namelist() else None
        extensible = ET.fromstring(archive.read("word/commentsExtensible.xml")) if "word/commentsExtensible.xml" in archive.namelist() else None
    text, starts, ends, position = [], {}, {}, 0
    for element in document.iter():
        value = ""
        if element.tag == q("w:t"):
            value = element.text or ""
        elif element.tag in (q("w:p"), q("w:br"), q("w:cr")):
            value = "\n"
        elif element.tag == q("w:tab"):
            value = "\t"
        text.append(value)
        position += len(value)
        if element.tag == q("w:commentRangeStart"):
            starts[element.get(q("w:id"))] = position
        elif element.tag == q("w:commentRangeEnd"):
            ends[element.get(q("w:id"))] = position
    text = "".join(text)
    for record in records:
        nodes = [c for c in comments if c.get(q("w:id")) == record["id"]]
        if len(nodes) != 1 or _text(nodes[0]) != record["text"] or nodes[0].get(q("w:author")) != record["author"]:
            fail(f"Comment {record['id']} did not survive verification.")
        paragraphs = nodes[0].findall("w:p", NS)
        if not paragraphs or paragraphs[-1].get(q("w14:paraId")) != record["para_id"]:
            fail(f"Paragraph ID for comment {record['id']} did not survive verification.")
        entries = [e for e in extended if e.get(q("w15:paraId")) == record["para_id"]]
        if len(entries) != 1 or entries[0].get(q("w15:paraIdParent"), "") != record["parent_para_id"]:
            fail(f"Thread linkage for comment {record['id']} did not survive verification.")
        for name in ("commentRangeStart", "commentRangeEnd", "commentReference"):
            if sum(e.get(q("w:id")) == record["id"] for e in document.iter(q(f"w:{name}"))) != 1:
                fail(f"Document anchor for comment {record['id']} did not survive verification.")
        start, end = starts[record["id"]], ends[record["id"]]
        if start > end or (start == end and record.get("anchor_text")):
            fail(f"Comment {record['id']} has an invalid text range.")
        if record.get("anchor_text") and _normalize(text[start:end]) != _normalize(record["anchor_text"]):
            fail(f"Anchored text for comment {record['id']} did not survive verification.")
        if record.get("parent_id") and (start, end) != (starts.get(record["parent_id"]), ends.get(record["parent_id"])):
            fail(f"Reply {record['id']} does not share its parent's anchor.")
        if record.get("durable_id"):
            ids = [] if identifiers is None else [e for e in identifiers if e.get(q("w16cid:paraId")) == record["para_id"]]
            if len(ids) != 1 or ids[0].get(q("w16cid:durableId")) != record["durable_id"]:
                fail(f"Durable ID for comment {record['id']} did not survive verification.")
            if extensible is not None and sum(e.get(q("w16cex:durableId")) == record["durable_id"] for e in extensible) != 1:
                fail(f"Extensible metadata for comment {record['id']} did not survive verification.")
