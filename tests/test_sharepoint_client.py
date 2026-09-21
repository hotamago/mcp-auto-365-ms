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


def test_parse_url_without_site_on_configured_host_falls_back_to_config():
    info = SharePointClient().parse_sharepoint_url("https://vingroupjsc.sharepoint.com/foo")
    assert info["site_path"] == "/sites/VF_AIDV"


def test_parse_url_on_a_foreign_host_never_borrows_the_configured_site():
    """The old code resolved *any* host without /sites/ against the config site."""
    info = SharePointClient().parse_sharepoint_url("https://contoso.sharepoint.com/foo")
    assert info["site_path"] == ""


@pytest.mark.parametrize(
    ("url", "site_path", "personal"),
    [
        ("https://t-my.sharepoint.com/personal/phuongnv24_vingroup_net/Documents/Microsoft Teams Chat Files/PSDK.zip",
         "/personal/phuongnv24_vingroup_net", True),
        ("https://t.sharepoint.com/teams/Ops/Shared%20Documents/a.xlsx", "/teams/Ops", False),
        ("https://t-my.sharepoint.com/:u:/g/personal/phuongnv24_vingroup_net/IQDZ", "/personal/phuongnv24_vingroup_net", True),
        ("https://t.sharepoint.com/:x:/r/sites/VF_AIDV/_layouts/15/Doc.aspx?sourcedoc={A}", "/sites/VF_AIDV", False),
    ],
)
def test_parse_url_recognises_every_site_kind(url, site_path, personal):
    info = SharePointClient().parse_sharepoint_url(url)
    assert info["site_path"] == site_path
    assert info["is_personal"] is personal


def test_strip_library_prefix_handles_onedrive_and_teams_sites():
    assert _strip_library_prefix("/personal/u_vingroup_net/Documents/A/b.zip") == "A/b.zip"
    assert _strip_library_prefix("/teams/Ops/Shared Documents/A") == "A"


# ------------------------------------------------------------ downloads


@pytest.fixture
def fetches(monkeypatch):
    """Record every byte fetch and the cookie host it was scoped to."""
    import sharepoint.client as spc

    calls: list[dict] = []
    client = SharePointClient()
    monkeypatch.setattr(
        client, "_cookie_headers", lambda accept="*/*", host="": {"Cookie": f"for:{host}", "Accept": accept}
    )
    monkeypatch.setattr(
        spc, "request_bytes", lambda url, headers=None, context="": calls.append({"url": url, "headers": headers}) or b"PK"
    )
    return client, calls


def test_onedrive_attachment_path_is_fetched_directly_with_its_own_host_cookie(fetches, tmp_path):
    client, calls = fetches
    url = "https://vingroupjsc-my.sharepoint.com/personal/phuongnv24_vingroup_net/Documents/Microsoft Teams Chat Files/PSDK.zip"
    client.download_link(url, str(tmp_path))
    assert len(calls) == 1
    assert calls[0]["url"].startswith("https://vingroupjsc-my.sharepoint.com/personal/")
    assert calls[0]["headers"]["Cookie"] == "for:vingroupjsc-my.sharepoint.com"
    assert (tmp_path / "PSDK.zip").read_bytes() == b"PK"


def test_generic_sharing_link_is_downloaded_with_download_flag(fetches, tmp_path):
    client, calls = fetches
    client.download_link("https://vingroupjsc-my.sharepoint.com/:u:/g/personal/u_vingroup_net/IQDZ", str(tmp_path))
    assert calls[0]["url"].endswith("?download=1")
    assert calls[0]["headers"]["Cookie"] == "for:vingroupjsc-my.sharepoint.com"


def test_item_url_prefers_graph_pre_authenticated_download():
    item = {"name": "a.xlsx", "@microsoft.graph.downloadUrl": "https://t.sharepoint.com/_layouts/15/download.aspx?tempauth=x"}
    assert SharePointClient()._item_file_url("drive", item).endswith("tempauth=x")


def test_item_url_fallback_uses_the_real_library_not_shared_documents(monkeypatch):
    """OneDrive's library is 'Documents'; hardcoding 'Shared Documents' broke it."""
    client = SharePointClient()
    monkeypatch.setattr(client, "_drive_web_url", lambda drive_id: "https://t-my.sharepoint.com/personal/u/Documents")
    item = {"name": "b c.zip", "parentReference": {"path": "/drives/x/root:/Teams Chat"}}
    assert client._item_file_url("drive", item) == "https://t-my.sharepoint.com/personal/u/Documents/Teams%20Chat/b%20c.zip"


def test_pre_authenticated_urls_are_fetched_without_cookies():
    headers = SharePointClient()._download_headers("https://t.sharepoint.com/_layouts/15/download.aspx?tempauth=abc")
    assert "Cookie" not in headers


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


def test_call_sharepoint_or_graph_prefers_cookie_channel(monkeypatch):
    client = SharePointClient()
    calls = []

    def fake_cookie_headers(accept="", host=""):
        return {"Cookie": f"rtFa=1; FedAuth=2; host={host}", "User-Agent": "test"}

    def fake_request_json(url, headers=None, method="GET", data=None, context=""):
        calls.append({"url": url, "headers": headers, "method": method})
        return {"id": "item123", "name": "doc.docx"}

    monkeypatch.setattr(client, "_cookie_headers", fake_cookie_headers)
    monkeypatch.setattr("sharepoint.client.request_json", fake_request_json)
    monkeypatch.setattr(client, "get_token", lambda *a, **k: pytest.fail("Azure CLI get_token must NOT be called when cookies work"))

    res = client.call_sharepoint_or_graph("/drives/drv1/root:/doc.docx", host="contoso.sharepoint.com")
    assert res["id"] == "item123"
    assert len(calls) == 1
    assert calls[0]["url"] == "https://contoso.sharepoint.com/_api/v2.0/drives/drv1/root:/doc.docx"
    assert "rtFa=1" in calls[0]["headers"]["Cookie"]


def test_call_sharepoint_or_graph_falls_back_to_azure_cli(monkeypatch):
    client = SharePointClient()
    calls = []

    def fake_cookie_headers(accept="", host=""):
        raise Mcp365Error("Cookie session missing")

    def fake_request_json(url, headers=None, method="GET", data=None, context=""):
        calls.append({"url": url, "headers": headers, "method": method})
        return {"id": "fallback_item"}

    monkeypatch.setattr(client, "_cookie_headers", fake_cookie_headers)
    monkeypatch.setattr(client, "get_token", lambda *a, **k: "mock-azure-cli-token")
    monkeypatch.setattr("sharepoint.client.request_json", fake_request_json)

    res = client.call_sharepoint_or_graph("/drives/drv1/items/it1")
    assert res["id"] == "fallback_item"
    assert len(calls) == 1
    assert calls[0]["url"] == "https://graph.microsoft.com/v1.0/drives/drv1/items/it1"
    assert calls[0]["headers"]["Authorization"] == "Bearer mock-azure-cli-token"


def test_put_file_bytes_cookie_channel(monkeypatch):
    client = SharePointClient()
    calls = []

    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": "rtFa=1; FedAuth=2"})
    monkeypatch.setattr(client, "_get_form_digest", lambda host="": "mock-digest-value")
    monkeypatch.setattr(
        "sharepoint.client.request_json",
        lambda url, headers=None, method="GET", data=None, context="": calls.append(
            {"url": url, "headers": headers, "method": method, "data": data}
        )
        or {"status": "ok"},
    )

    res = client.put_file_bytes("drv1", "it1", b"new content", if_match="W/'123'")
    assert res["status"] == "ok"
    assert len(calls) == 1
    assert calls[0]["method"] == "PUT"
    assert calls[0]["headers"]["If-Match"] == "W/'123'"
    assert calls[0]["headers"]["X-RequestDigest"] == "mock-digest-value"
    assert calls[0]["data"] == b"new content"


def test_get_form_digest_caching(monkeypatch):
    client = SharePointClient()
    digest_calls = []

    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": "c=1"})
    monkeypatch.setattr(
        "sharepoint.client.request_json",
        lambda url, headers=None, method="GET", data=None, context="": digest_calls.append(url)
        or {"d": {"GetContextWebInformation": {"FormDigestValue": "digest_abc", "FormDigestTimeoutSeconds": 1800}}},
    )

    d1 = client._get_form_digest("contoso.sharepoint.com")
    d2 = client._get_form_digest("contoso.sharepoint.com")
    assert d1 == "digest_abc"
    assert d2 == "digest_abc"
    assert len(digest_calls) == 1


def test_probe_graph_reports_ok_when_cookie_channel_is_active(monkeypatch):
    from common import health
    from common.errors import Mcp365Error

    def mock_get_token(self, *a, **k):
        raise Mcp365Error("Azure CLI không phản hồi sau 60s.")

    monkeypatch.setattr(SharePointClient, "get_token", mock_get_token)
    monkeypatch.setattr(
        "common.chrome_cookies.ChromeCookieDecryptor.get_cookies_for_domain",
        lambda domain, names: {"rtFa": "mock-rtfa", "FedAuth": "mock-fedauth"},
    )

    probe = health._probe_graph()
    assert probe["status"] == "✅"
    assert "kênh phụ" in probe["title"]
    assert "cookie Chrome" in probe["detail"]
