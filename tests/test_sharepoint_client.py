"""SharePoint helpers verifiable without a network."""

from __future__ import annotations

import io
import zipfile

import pytest

from common.errors import Mcp365Error
from sharepoint.client import SharePointClient, _strip_library_prefix, human_size


def _docx(paragraphs: list[str]) -> bytes:
    body = "".join(f"<w:p><w:r><w:t>{p}</w:t></w:r></w:p>" for p in paragraphs)
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("word/document.xml", f"<w:document><w:body>{body}</w:body></w:document>")
    return buf.getvalue()


def _xlsx(sheets: list[str]) -> bytes:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for name in sheets:
            zf.writestr(f"xl/worksheets/{name}", "<worksheet/>")
    return buf.getvalue()


def test_human_size():
    assert human_size(512) == "512 B"
    assert human_size(2048) == "2.0 KB"
    assert human_size(5 * 1024 * 1024) == "5.00 MB"


def test_strip_library_prefix():
    assert _strip_library_prefix("/sites/VF_AIDV/Shared Documents/A/B") == "A/B"
    assert _strip_library_prefix("A/B") == "A/B"
    assert _strip_library_prefix("") == ""


def test_extract_lines_from_docx():
    lines = SharePointClient.extract_lines(_docx(["Tiêu đề", "", "Nội dung"]), "a.docx")
    assert lines == ["Tiêu đề", "Nội dung"]


def test_extract_lines_from_text():
    assert SharePointClient.extract_lines(b"a\nb\n", "notes.md") == ["a", "b"]


def test_extract_lines_from_xlsx_lists_sheets():
    assert SharePointClient.extract_lines(_xlsx(["sheet2.xml", "sheet1.xml"]), "b.xlsx") == [
        "Sheet: sheet1.xml",
        "Sheet: sheet2.xml",
    ]


def test_extract_lines_returns_none_for_binary():
    assert SharePointClient.extract_lines(b"\x00\x01binary", "image.png") is None


def test_corrupt_docx_does_not_raise():
    assert SharePointClient.extract_lines(b"not a zip", "broken.docx") is None


def test_render_diff_reports_additions_and_deletions():
    out = "\n".join(SharePointClient._render_diff(["a", "b"], ["a", "c"], "A", "B", 10, 10))
    assert "+1 thêm" in out and "-1 xoá" in out


def test_render_diff_reports_identical():
    out = "\n".join(SharePointClient._render_diff(["a"], ["a"], "A", "B", 1, 1))
    assert "Không có khác biệt" in out


def test_render_diff_falls_back_to_size_for_binary():
    out = "\n".join(SharePointClient._render_diff(None, None, "A", "B", 100, 150))
    assert "+50 bytes" in out


def test_parse_url_extracts_site_and_sourcedoc():
    info = SharePointClient().parse_sharepoint_url(
        "https://contoso.sharepoint.com/sites/Eng/_layouts/15/Doc.aspx?sourcedoc={ABC-123}&file=x.docx"
    )
    assert info["site_path"] == "/sites/Eng"
    assert info["sourcedoc"] == "ABC-123"
    assert info["type"] == "document"
    assert info["file_name"] == "x.docx"


def test_parse_url_without_site_falls_back_to_config():
    info = SharePointClient().parse_sharepoint_url("https://contoso.sharepoint.com/foo")
    assert info["site_path"] == "/sites/VF_AIDV"


def test_download_rejects_garbage_input():
    with pytest.raises(Mcp365Error):
        SharePointClient().download_link("khong-phai-link")


def test_sync_folder_requires_existing_directory():
    with pytest.raises(Mcp365Error) as excinfo:
        SharePointClient().sync_folder_up("/khong/ton/tai", "Target")
    assert excinfo.value.remediation


def test_sync_dry_run_never_uploads(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_text("hello")
    client = SharePointClient()
    monkeypatch.setattr(client, "resolve_drive", lambda url="": ({}, "drive1"))
    monkeypatch.setattr(client, "list_folder_contents", lambda drive, path: [])
    uploads = []
    monkeypatch.setattr(client, "upload_file", lambda *a, **k: uploads.append(a))

    report = client.sync_folder_up(str(tmp_path), "Target", dry_run=True)
    assert uploads == []
    assert "NEW" in report and "a.txt" in report


def test_sync_skips_files_already_newer_remotely(tmp_path, monkeypatch):
    (tmp_path / "a.txt").write_text("hello")
    client = SharePointClient()
    monkeypatch.setattr(client, "resolve_drive", lambda url="": ({}, "drive1"))
    monkeypatch.setattr(
        client,
        "list_folder_contents",
        lambda drive, path: [{"name": "a.txt", "lastModifiedDateTime": "2099-01-01T00:00:00Z"}],
    )
    assert "không có file nào" in client.sync_folder_up(str(tmp_path), "Target", dry_run=True).lower()
