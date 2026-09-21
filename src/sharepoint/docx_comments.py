"""Add review comments to a ``.docx`` in memory.

Pure function over bytes: no network. Each comment is anchored to the first
paragraph whose text contains the given ``anchor`` phrase - on the exact run
when the phrase sits inside one, otherwise on the whole paragraph (Word splits
text across runs for spell-check and revision marks, so a visible phrase is
often not contiguous in the XML).

lxml is used instead of ElementTree on purpose: ElementTree drops namespace
declarations it considers unused, and Word refuses a file whose
``mc:Ignorable`` names a prefix that is no longer declared.
"""

from __future__ import annotations

import io
import re
import zipfile
from datetime import UTC, datetime
from typing import Any

from lxml import etree

from common.errors import Mcp365Error

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_R = "http://schemas.openxmlformats.org/package/2006/relationships"
_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
_COMMENTS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
_COMMENTS_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"


def _w(tag: str) -> str:
    return f"{{{W}}}{tag}"


def _norm(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip().casefold()


def _text(el: etree._Element) -> str:
    return "".join(t.text or "" for t in el.iter(_w("t")))


def _xml(el: etree._Element) -> bytes:
    return etree.tostring(el, xml_declaration=True, encoding="UTF-8", standalone=True)


def _find_paragraph(body: etree._Element, anchor: str) -> etree._Element | None:
    """Innermost paragraph containing ``anchor`` (text boxes nest paragraphs)."""
    wanted = _norm(anchor)
    hits = [p for p in body.iter(_w("p")) if wanted in _norm(_text(p))]
    for p in hits:
        if not any(wanted in _norm(_text(inner)) for inner in p.iter(_w("p")) if inner is not p):
            return p
    return None


def _comment_element(cid: int, author: str, initials: str, text: str, when: str) -> etree._Element:
    comment = etree.Element(_w("comment"))
    comment.set(_w("id"), str(cid))
    comment.set(_w("author"), author)
    comment.set(_w("date"), when)
    comment.set(_w("initials"), initials)
    for line in text.splitlines() or [""]:
        p = etree.SubElement(comment, _w("p"))
        r = etree.SubElement(p, _w("r"))
        t = etree.SubElement(r, _w("t"))
        t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = line
    return comment


def _mark(p: etree._Element, anchor: str, cid: int) -> str:
    """Insert range start/end and the reference run; return where it landed."""
    start, end = etree.Element(_w("commentRangeStart")), etree.Element(_w("commentRangeEnd"))
    for el in (start, end):
        el.set(_w("id"), str(cid))
    ref_run = etree.Element(_w("r"))
    etree.SubElement(ref_run, _w("commentReference")).set(_w("id"), str(cid))

    wanted = _norm(anchor)
    run = next((r for r in p.iter(_w("r")) if wanted in _norm(_text(r))), None)
    if run is not None:
        run.addprevious(start)
        run.addnext(end)
        end.addnext(ref_run)
        return "đúng cụm từ"

    ppr = p.find(_w("pPr"))
    p.insert(p.index(ppr) + 1 if ppr is not None else 0, start)
    p.append(end)
    p.append(ref_run)
    return "cả đoạn"


def add_comments(data: bytes, comments: list[dict[str, str]], author: str, initials: str = "") -> tuple[bytes, list[dict[str, Any]]]:
    """Return the new ``.docx`` bytes and where each comment was anchored.

    Raises if the file is not a Word document or any anchor cannot be found -
    a comment that silently lands nowhere is worse than none.
    """
    if not comments:
        raise Mcp365Error("Không có comment nào để thêm.", 'Truyền comments=[{"anchor": "...", "text": "..."}].')
    try:
        zin = zipfile.ZipFile(io.BytesIO(data))
        doc = etree.fromstring(zin.read("word/document.xml"))
    except (zipfile.BadZipFile, KeyError, etree.XMLSyntaxError) as exc:
        raise Mcp365Error(
            f"File không phải tài liệu Word hợp lệ: {exc}",
            "Bản tải về có thể là trang HTML (link chia sẻ/đăng nhập) chứ không phải .docx. "
            "Tải lại bằng đường dẫn trực tiếp của file.",
        ) from exc

    names = zin.namelist()
    if "word/comments.xml" in names:
        comments_root = etree.fromstring(zin.read("word/comments.xml"))
    else:
        comments_root = etree.Element(_w("comments"), nsmap={"w": W})
    used = [int(x) for x in doc.xpath("//w:commentRangeStart/@w:id | //w:commentReference/@w:id", namespaces={"w": W})]
    used += [int(x) for x in comments_root.xpath("//w:comment/@w:id", namespaces={"w": W})]
    next_id = max(used, default=-1) + 1

    body = doc.find(_w("body"))
    missing = [c["anchor"] for c in comments if _find_paragraph(body, c.get("anchor", "")) is None]
    if missing:
        raise Mcp365Error(
            "Không tìm thấy đoạn nào chứa: " + "; ".join(f"'{a}'" for a in missing),
            "Chép nguyên văn một cụm từ ngắn (3-8 chữ) có trong tài liệu làm anchor.",
        )

    initials = initials or "".join(w[0] for w in author.split("(")[0].split() if w)[:3].upper()
    when = datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    report = []
    for c in comments:
        p = _find_paragraph(body, c["anchor"])
        placed = _mark(p, c["anchor"], next_id)
        comments_root.append(_comment_element(next_id, author, initials, c["text"], when))
        report.append({"id": next_id, "anchor": c["anchor"], "text": c["text"], "placed": placed, "context": _text(p)[:120]})
        next_id += 1

    rels = etree.fromstring(zin.read("word/_rels/document.xml.rels"))
    if not any(r.get("Type") == _COMMENTS_REL for r in rels):
        ids = [int(m.group(1)) for r in rels if (m := re.match(r"rId(\d+)$", r.get("Id", "")))]
        rel = etree.SubElement(rels, f"{{{_R}}}Relationship")
        rel.set("Id", f"rId{max(ids, default=0) + 1}")
        rel.set("Type", _COMMENTS_REL)
        rel.set("Target", "comments.xml")

    types = etree.fromstring(zin.read("[Content_Types].xml"))
    if not any(o.get("PartName") == "/word/comments.xml" for o in types.iter(f"{{{_CT}}}Override")):
        override = etree.SubElement(types, f"{{{_CT}}}Override")
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", _COMMENTS_CT)

    replaced = {
        "word/document.xml": _xml(doc),
        "word/comments.xml": _xml(comments_root),
        "word/_rels/document.xml.rels": _xml(rels),
        "[Content_Types].xml": _xml(types),
    }
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            zout.writestr(info, replaced.pop(info.filename, None) or zin.read(info.filename))
        for name, payload in replaced.items():  # parts that did not exist before
            zout.writestr(name, payload)
    return out.getvalue(), report


def render_report(report: list[dict[str, Any]]) -> str:
    rows = [f"{len(report)} comment:\n", "| # | Gắn vào | Nội dung comment |", "| --- | --- | --- |"]
    for r in report:
        where = f"“{r['anchor']}” ({r['placed']})"
        rows.append(f"| {r['id']} | {where} | {r['text'].replace(chr(10), ' ')} |")
    return "\n".join(rows)
