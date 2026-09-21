"""Anchoring comments inside a .docx built in memory."""

from __future__ import annotations

import io
import zipfile

import pytest

from common.errors import Mcp365Error
from sharepoint.docx_comments import add_comments

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"


def _docx(body: str) -> bytes:
    """Minimal Word package. mc:Ignorable names w14 on purpose: dropping that
    declaration on save is what makes Word call a file corrupt."""
    doc = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{W}" xmlns:mc="{MC}" '
        'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" mc:Ignorable="w14">'
        f"<w:body>{body}</w:body></w:document>"
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as z:
        z.writestr(
            "[Content_Types].xml",
            '<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
            '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>',
        )
        z.writestr(
            "word/_rels/document.xml.rels",
            '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://x/styles" Target="styles.xml"/></Relationships>',
        )
        z.writestr("word/document.xml", doc)
    return buf.getvalue()


def _p(*runs: str) -> str:
    from xml.sax.saxutils import escape

    return "<w:p><w:pPr/>" + "".join(f"<w:r><w:t>{escape(r)}</w:t></w:r>" for r in runs) + "</w:p>"


BODY = _p("Step 3: tối đa 1 lần nếu trip < 20 phút.") + _p("ttl = ", "null", "; huỷ khi ON_END_TRIP")


def _parts(data: bytes) -> dict[str, str]:
    with zipfile.ZipFile(io.BytesIO(data)) as z:
        return {n: z.read(n).decode("utf-8") for n in z.namelist()}


def test_comment_anchors_on_the_run_and_wires_the_package():
    out, report = add_comments(_docx(BODY), [{"anchor": "trip < 20 phút", "text": "Trip chưa kết thúc thì sao biết?"}], "Nguyễn Hoàng Sơn")
    parts = _parts(out)
    doc = parts["word/document.xml"]
    assert '<w:commentRangeStart w:id="0"/>' in doc
    assert '<w:commentRangeEnd w:id="0"/>' in doc
    assert '<w:commentReference w:id="0"/>' in doc
    assert "Trip chưa kết thúc thì sao biết?" in parts["word/comments.xml"]
    assert 'w:author="Nguyễn Hoàng Sơn"' in parts["word/comments.xml"]
    assert "relationships/comments" in parts["word/_rels/document.xml.rels"]
    assert "/word/comments.xml" in parts["[Content_Types].xml"]
    assert report[0]["placed"] == "đúng cụm từ"


def test_namespace_declarations_survive_the_round_trip():
    out, _ = add_comments(_docx(BODY), [{"anchor": "ON_END_TRIP", "text": "x"}], "A")
    doc = _parts(out)["word/document.xml"]
    assert 'mc:Ignorable="w14"' in doc
    assert "xmlns:w14=" in doc


def test_phrase_split_across_runs_falls_back_to_the_paragraph():
    """Word splits text into runs; 'ttl = null' spans two of them here."""
    out, report = add_comments(_docx(BODY), [{"anchor": "ttl = null", "text": "Treo vô hạn khi xe offline?"}], "A")
    assert report[0]["placed"] == "cả đoạn"
    assert '<w:commentReference w:id="0"/>' in _parts(out)["word/document.xml"]


def test_several_comments_get_distinct_ids():
    out, report = add_comments(
        _docx(BODY), [{"anchor": "trip < 20", "text": "a"}, {"anchor": "ON_END_TRIP", "text": "b"}], "A"
    )
    assert [r["id"] for r in report] == [0, 1]
    assert _parts(out)["word/comments.xml"].count("<w:comment ") == 2


def test_missing_anchor_refuses_the_whole_batch():
    with pytest.raises(Mcp365Error) as excinfo:
        add_comments(_docx(BODY), [{"anchor": "trip < 20", "text": "a"}, {"anchor": "không có", "text": "b"}], "A")
    assert "'không có'" in excinfo.value.message


def test_html_instead_of_docx_is_explained():
    with pytest.raises(Mcp365Error) as excinfo:
        add_comments(b"<html>sign in</html>", [{"anchor": "x", "text": "y"}], "A")
    assert "HTML" in excinfo.value.remediation
