"""SharePoint helpers verifiable without a network."""

from __future__ import annotations

import io
import urllib.parse
import zipfile
from pathlib import Path

import pytest

from common.errors import AuthExpiredError, CAEChallengeError, ConcurrentEditError, Mcp365Error
from sharepoint.client import SharePointClient, _both_channels_failed, _strip_library_prefix, human_size


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

    def to_file(url, dest, headers=None, context="", **kwargs):
        calls.append({"url": url, "headers": headers})
        Path(dest).write_bytes(b"PK")
        return 2

    monkeypatch.setattr(spc, "request_to_file", to_file)
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


# The link from the 22/09 report: a ``/:x:/r/`` wrapper around Doc.aspx whose
# ``file=`` is a stale name - the item is really called VSDK.xlsx.
_DOC_LINK = (
    "https://vingroupjsc.sharepoint.com/:x:/r/sites/VF_AIDV/_layouts/15/Doc.aspx?"
    "sourcedoc=%7B96E95EB2-6E7A-4757-B4E8-A0B4D2178947%7D&file=VinFast-IVI-SDK-Components-1.0.3.xlsx"
    "&action=default&mobileredirect=true&wdwpf=doclib-c"
)
_GUID = "96E95EB2-6E7A-4757-B4E8-A0B4D2178947"
_ITEM = {
    "id": "01ABC",
    "name": "VSDK.xlsx",
    "file": {"mimeType": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"},
    "parentReference": {"driveId": "b!drive", "path": "/drives/b!drive/root:/S5/02. Technical Docs"},
    "@content.downloadUrl": "https://vingroupjsc.sharepoint.com/sites/VF_AIDV/_layouts/15/download.aspx?UniqueId=x&tempauth=t",
}
_VIEWER_PAGE = b"<!DOCTYPE html><html><head><title>Excel</title></head><body>WacFrame</body></html>"


@pytest.mark.parametrize("kind", ["x", "w", "p"])
def test_doc_aspx_sharing_link_is_parsed_to_its_guid(kind):
    info = SharePointClient().parse_sharepoint_url(_DOC_LINK.replace("/:x:/", f"/:{kind}:/"))
    assert info["sourcedoc"] == _GUID
    assert info["type"] == "document"
    assert info["site_path"] == "/sites/VF_AIDV"
    assert info["file_name"] == "VinFast-IVI-SDK-Components-1.0.3.xlsx"  # stale label, must not name the file


@pytest.fixture
def guid_lookup(monkeypatch):
    """Resolve any link to ``_ITEM`` without the network, recording the GUIDs asked for."""
    asked: list[str] = []

    def resolve_drive(self, url=""):
        return self.parse_sharepoint_url(url) if url else {}, "b!drive"

    monkeypatch.setattr(SharePointClient, "resolve_drive", resolve_drive)
    monkeypatch.setattr(SharePointClient, "get_item_by_guid", lambda self, drive, guid: asked.append(guid) or dict(_ITEM))
    return asked


def test_doc_aspx_link_downloads_the_item_bytes_under_its_real_name(fetches, guid_lookup, tmp_path):
    client, calls = fetches
    report = client.download_link(_DOC_LINK, str(tmp_path))

    assert guid_lookup == [_GUID]
    assert [c["url"] for c in calls] == [_ITEM["@content.downloadUrl"]]  # never Doc.aspx...&download=1
    assert "Cookie" not in calls[0]["headers"]  # pre-authenticated URL
    assert [p.name for p in tmp_path.iterdir()] == ["VSDK.xlsx"]
    assert "Downloaded 1/1" in report and "VSDK.xlsx" in report


@pytest.fixture
def served(monkeypatch):
    """``request_to_file`` fake: ``pages[url]`` is the body; HTML is refused like the real one."""
    import sharepoint.client as spc
    from common.http import looks_like_html

    pages: dict[str, bytes] = {}
    fetched: list[str] = []

    def to_file(url, dest, headers=None, context="", reject_html=False, **kwargs):
        fetched.append(url)
        body = pages[url]
        if reject_html and looks_like_html("", body):
            raise spc.HtmlPageError("html", "")
        Path(dest).write_bytes(body)
        return len(body)

    monkeypatch.setattr(spc, "request_to_file", to_file)
    client = SharePointClient()
    monkeypatch.setattr(client, "_cookie_headers", lambda accept="*/*", host="": {"Cookie": f"for:{host}"})
    return client, pages, fetched


def test_html_answer_moves_on_to_the_next_url(served, tmp_path):
    client, pages, fetched = served
    client._drive_cache["web:b!drive"] = "https://vingroupjsc.sharepoint.com/sites/VF_AIDV/Shared%20Documents"
    path_url = "https://vingroupjsc.sharepoint.com/sites/VF_AIDV/Shared%20Documents/S5/02.%20Technical%20Docs/VSDK.xlsx"
    pages[_ITEM["@content.downloadUrl"]] = _VIEWER_PAGE
    pages[path_url] = b"PK\x03\x04real"

    results: list = []
    client._save_item("b!drive", dict(_ITEM), tmp_path, results)

    assert fetched == [_ITEM["@content.downloadUrl"], path_url]
    assert (tmp_path / "VSDK.xlsx").read_bytes() == b"PK\x03\x04real"


def test_only_html_available_fails_loudly_and_writes_nothing(served, guid_lookup, tmp_path):
    from common.errors import HtmlPageError

    client, pages, _fetched = served
    client._drive_cache["web:b!drive"] = "https://vingroupjsc.sharepoint.com/sites/VF_AIDV/Shared%20Documents"
    pages[_ITEM["@content.downloadUrl"]] = _VIEWER_PAGE
    pages["https://vingroupjsc.sharepoint.com/sites/VF_AIDV/Shared%20Documents/S5/02.%20Technical%20Docs/VSDK.xlsx"] = (
        _VIEWER_PAGE
    )
    with pytest.raises(HtmlPageError) as excinfo:
        client.download_link(_DOC_LINK, str(tmp_path))
    assert "VSDK.xlsx" in excinfo.value.message
    assert "GUID" in excinfo.value.remediation
    assert list(tmp_path.iterdir()) == []


def test_sharing_link_that_answers_with_a_page_is_resolved_through_shares(served, monkeypatch, tmp_path):
    """A short Office link has no GUID; download=1 yields the viewer, /shares yields the item."""
    client, pages, fetched = served
    link = "https://vingroupjsc.sharepoint.com/:x:/g/sites/VF_AIDV/EaBcDeF?e=x1"
    pages[f"{link}&download=1"] = _VIEWER_PAGE
    pages[_ITEM["@content.downloadUrl"]] = b"PK\x03\x04real"
    monkeypatch.setattr(client, "_resolve_shared_item", lambda url: ("b!drive", dict(_ITEM)))

    client.download_link(link, str(tmp_path))

    assert fetched == [f"{link}&download=1", _ITEM["@content.downloadUrl"]]
    assert [p.name for p in tmp_path.iterdir()] == ["VSDK.xlsx"]


# The 26/09 case: a Teams chat attachment ``WMC_UC01_005.yaml``. ``.yaml`` was
# not in the extension whitelist, so the file was opened as a folder and its
# ``/children`` listed: 422 getChildrenOnNonFolder on cookie, 403 on Graph.
_CHAT_FILES = "https://vingroupjsc-my.sharepoint.com/personal/hiennt_vingroup_net/Documents/Microsoft%20Teams%20Chat%20Files"


@pytest.fixture
def graph(monkeypatch):
    """Fake Graph for the folder branch: ``items[endpoint]`` answers; every endpoint asked is recorded."""
    items: dict[str, object] = {}
    asked: list[str] = []

    def resolve_drive(self, url=""):
        return self.parse_sharepoint_url(url), "b!drive"

    def call_graph(self, path, method="GET", body=None, context=""):
        asked.append(path)
        answer = items[path]
        if isinstance(answer, Exception):
            raise answer
        return answer

    monkeypatch.setattr(SharePointClient, "resolve_drive", resolve_drive)
    monkeypatch.setattr(SharePointClient, "call_graph", call_graph)
    return items, asked


@pytest.mark.parametrize(
    "path, is_file",
    [
        ("/personal/u/Documents/Chat Files/WMC_UC01_005.yaml", True),
        ("/personal/u/Documents/Chat Files/a.yml", True),
        ("/teams/Ops/Shared Documents/logs.tar.gz", True),
        ("/sites/S/Shared Documents/app.7z", True),
        ("/sites/S/Shared Documents/Chat Files", False),
        ("/sites/S/Shared Documents/S5/02. Technical Docs", False),
        ("/sites/S/Shared Documents/Release/v1.2", False),
        ("/sites/S/Shared Documents/Forms/AllItems.aspx", False),
        ("/sites/S/_layouts/15/Doc.aspx", False),
    ],
)
def test_file_name_shape_does_not_depend_on_a_whitelist(path, is_file):
    from sharepoint.client import _looks_like_file

    assert _looks_like_file(path) is is_file


@pytest.mark.parametrize("name", ["WMC_UC01_005.yaml", "config.yml", "Main.kt", "run.log", "app-release.apk"])
def test_onedrive_file_with_any_extension_is_fetched_directly(fetches, monkeypatch, tmp_path, name):
    client, calls = fetches
    monkeypatch.setattr(SharePointClient, "resolve_drive", lambda self, url="": pytest.fail("a file is not a folder"))

    report = client.download_link(f"{_CHAT_FILES}/{name}", str(tmp_path))

    assert [c["url"] for c in calls] == [f"{_CHAT_FILES}/{name}"]
    assert calls[0]["headers"]["Cookie"] == "for:vingroupjsc-my.sharepoint.com"
    assert (tmp_path / name).read_bytes() == b"PK"
    assert "Downloaded 1/1" in report


@pytest.mark.parametrize(
    "url, rel",
    [
        (_CHAT_FILES, "Microsoft%20Teams%20Chat%20Files"),
        (f"{_CHAT_FILES}/02.%20Specs", "Microsoft%20Teams%20Chat%20Files/02.%20Specs"),
        (
            "https://vingroupjsc.sharepoint.com/sites/VF_AIDV/Shared%20Documents/Forms/AllItems.aspx"
            "?id=%2Fsites%2FVF_AIDV%2FShared%20Documents%2FS5",
            "S5",
        ),
    ],
)
def test_folder_url_still_lists_the_folder(fetches, graph, tmp_path, url, rel):
    client, calls = fetches
    items, asked = graph
    items[f"/drives/b!drive/root:/{rel}"] = {"id": "F1", "name": "folder", "folder": {"childCount": 0}}
    items["/drives/b!drive/items/F1/children"] = {"value": []}

    client.download_link(url, str(tmp_path))

    assert calls == []  # no direct fetch of a folder path or a folder-view page
    assert asked == [f"/drives/b!drive/root:/{rel}", "/drives/b!drive/items/F1/children"]


def test_graph_item_that_is_a_file_is_downloaded_not_listed(fetches, graph, tmp_path):
    """No extension, so no direct fetch: Graph says it is a file, and ``/children`` is never asked."""
    client, calls = fetches
    items, asked = graph
    download_url = "https://vingroupjsc-my.sharepoint.com/personal/u/_layouts/15/download.aspx?UniqueId=x&tempauth=t"
    items["/drives/b!drive/root:/Microsoft%20Teams%20Chat%20Files/Makefile"] = {
        "id": "01F",
        "name": "Makefile",
        "file": {"mimeType": "application/octet-stream"},
        "@content.downloadUrl": download_url,
    }

    client.download_link(f"{_CHAT_FILES}/Makefile", str(tmp_path))

    assert not any(endpoint.endswith("/children") for endpoint in asked)
    assert [c["url"] for c in calls] == [download_url]
    assert (tmp_path / "Makefile").read_bytes() == b"PK"


def test_refused_direct_fetch_falls_back_to_the_graph_item(served, graph, tmp_path):
    """A sign-in page instead of the .yaml: Graph finds the item and its pre-authenticated URL delivers it."""
    client, pages, fetched = served
    items, asked = graph
    direct = f"{_CHAT_FILES}/WMC_UC01_005.yaml"
    pages[direct] = _VIEWER_PAGE
    pages[_ITEM["@content.downloadUrl"]] = b"uc: 01\n"
    items["/drives/b!drive/root:/Microsoft%20Teams%20Chat%20Files/WMC_UC01_005.yaml"] = {
        **_ITEM,
        "name": "WMC_UC01_005.yaml",
    }

    client.download_link(direct, str(tmp_path))

    assert fetched == [direct, _ITEM["@content.downloadUrl"]]
    assert not any(endpoint.endswith("/children") for endpoint in asked)
    assert (tmp_path / "WMC_UC01_005.yaml").read_bytes() == b"uc: 01\n"


def test_folder_with_a_dot_in_its_name_falls_back_to_the_folder_branch(served, graph, tmp_path):
    client, pages, fetched = served
    items, asked = graph
    direct = f"{_CHAT_FILES}/com.vinfast.app"
    pages[direct] = _VIEWER_PAGE  # the folder view page, refused before anything is written
    items["/drives/b!drive/root:/Microsoft%20Teams%20Chat%20Files/com.vinfast.app"] = {"id": "F2", "folder": {}}
    items["/drives/b!drive/items/F2/children"] = {"value": []}

    client.download_link(direct, str(tmp_path))

    assert fetched == [direct]
    assert asked[-1] == "/drives/b!drive/items/F2/children"
    assert list(tmp_path.iterdir()) == []


def test_when_both_routes_fail_both_errors_are_reported(served, graph, tmp_path):
    client, pages, _fetched = served
    items, _asked = graph
    direct = f"{_CHAT_FILES}/WMC_UC01_005.yaml"
    pages[direct] = _VIEWER_PAGE
    items["/drives/b!drive/root:/Microsoft%20Teams%20Chat%20Files/WMC_UC01_005.yaml"] = Mcp365Error("HTTP 404 itemNotFound")

    with pytest.raises(Mcp365Error) as excinfo:
        client.download_link(direct, str(tmp_path))

    assert "trang web (HTML" in excinfo.value.message
    assert "HTTP 404 itemNotFound" in excinfo.value.message


def test_listing_the_children_of_a_file_is_reported_as_a_link_error_not_a_permission_one(graph, tmp_path):
    items, _asked = graph
    items["/drives/b!drive/items/01F/children"] = _both_channels_failed(
        "liệt kê nội dung thư mục",
        Mcp365Error('HTTP 422 Unprocessable Entity.\nPhản hồi: {"error":{"code":"getChildrenOnNonFolder"}}'),
        AuthExpiredError("Bị từ chối truy cập (HTTP 403).", "Kiểm tra bạn có quyền trên site/tài nguyên này"),
    )

    with pytest.raises(Mcp365Error) as excinfo:
        SharePointClient()._sync_folder_down("b!drive", "01F", tmp_path, [])

    assert "nhận diện link" in excinfo.value.message
    assert "KHÔNG phải lỗi thiếu quyền" in excinfo.value.message
    assert "403" not in excinfo.value.message
    assert "quyền trên site" not in excinfo.value.remediation


def test_share_id_is_unpadded_base64url():
    from sharepoint.client import _share_id

    assert _share_id("https://t.sharepoint.com/:x:/g/a") == "u!aHR0cHM6Ly90LnNoYXJlcG9pbnQuY29tLzp4Oi9nL2E"


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


def test_cookie_write_carries_the_digest_and_extra_headers(monkeypatch):
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

    res = client.call_sharepoint_or_graph(
        "/drives/drv1/items/it1/content", method="PUT", data=b"new content", extra_headers={"If-Match": "W/'123'"}
    )
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

    d1 = client._get_form_digest("https://contoso.sharepoint.com/sites/Eng")
    d2 = client._get_form_digest("https://contoso.sharepoint.com/sites/Eng/")
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


# ------------------------------------------------- errors across channels


def test_both_channels_failed_shows_each_error_and_each_fix():
    err = _both_channels_failed(
        "tạo thư mục 'S5'",
        AuthExpiredError("Bị từ chối truy cập (HTTP 403).", "Kiểm tra quyền trên site."),
        CAEChallengeError("CAE đã thu hồi token (HTTP 401).", "Chạy: az login"),
    )
    assert "tạo thư mục 'S5'" in err.message
    assert "HTTP 403" in err.message and "HTTP 401" in err.message
    assert "Kiểm tra quyền trên site." in err.remediation and "Chạy: az login" in err.remediation


def test_real_cookie_error_is_not_hidden_by_the_graph_retry(monkeypatch):
    """SharePoint's own 403 used to be replaced by Graph's CAE 401."""
    client = SharePointClient()
    graph_calls = []

    def fake_request_json(url, headers=None, method="GET", data=None, context=""):
        if url.startswith("https://graph.microsoft.com"):
            graph_calls.append(url)
            raise CAEChallengeError("CAE đã thu hồi token (HTTP 401).", "Chạy: az login")
        raise AuthExpiredError("Bị từ chối truy cập (HTTP 403).", "Kiểm tra quyền trên site.")

    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": "c"})
    monkeypatch.setattr(client, "get_token", lambda *a, **k: "tok")
    monkeypatch.setattr("sharepoint.client.request_json", fake_request_json)

    with pytest.raises(Mcp365Error) as excinfo:
        client.call_sharepoint_or_graph("/drives/d/root:/A", context="kiểm tra thư mục 'A'")
    text = str(excinfo.value)
    assert len(graph_calls) == 2  # the token-refresh retry also ends up in the combined report
    assert "HTTP 403" in text and "HTTP 401" in text
    assert "Kiểm tra quyền trên site." in text and "az login" in text


def test_concurrent_edit_on_the_cookie_channel_does_not_fall_back(monkeypatch):
    client = SharePointClient()

    def fake_request_json(url, headers=None, method="GET", data=None, context=""):
        if url.endswith("/_api/contextinfo"):
            return {"d": {"GetContextWebInformation": {"FormDigestValue": "dg"}}}
        raise ConcurrentEditError("File đã bị người khác sửa (HTTP 412). Chưa ghi gì cả.", "Không ghi đè.")

    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": "c"})
    monkeypatch.setattr(client, "get_token", lambda *a, **k: pytest.fail("a refused write must not be retried on Graph"))
    monkeypatch.setattr("sharepoint.client.request_json", fake_request_json)

    with pytest.raises(ConcurrentEditError):
        client.call_sharepoint_or_graph(
            "/drives/drv1/items/it1/content", method="PUT", data=b"x", extra_headers={"If-Match": '"etag"'}
        )


def test_resolve_drive_reports_the_cookie_error_when_graph_also_fails(monkeypatch):
    client = SharePointClient()

    def refused(url, headers=None, method="GET", data=None, context=""):
        raise AuthExpiredError("Bị từ chối truy cập (HTTP 403).", "Kiểm tra quyền trên site.")

    def cae(hostname, site_path):
        raise CAEChallengeError("CAE đã thu hồi token (HTTP 401).", "Chạy: az login")

    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": "c"})
    monkeypatch.setattr("sharepoint.client.request_json", refused)
    monkeypatch.setattr(client, "get_site_id", cae)

    with pytest.raises(Mcp365Error) as excinfo:
        client.resolve_drive()
    assert "HTTP 403" in str(excinfo.value) and "HTTP 401" in str(excinfo.value)


# ------------------------------------------------------ site-scoped writes


def _digest_response(value: str = "site-digest") -> dict:
    return {"d": {"GetContextWebInformation": {"FormDigestValue": value, "FormDigestTimeoutSeconds": 1800}}}


def test_digest_is_requested_from_the_site_not_the_host_root(monkeypatch):
    """A digest from https://host/_api/contextinfo is refused by /sites/VF_AIDV."""
    from common.config import get_config

    client = SharePointClient()
    urls = []
    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": f"for:{host}"})
    monkeypatch.setattr(
        "sharepoint.client.request_json",
        lambda url, headers=None, method="GET", data=None, context="": urls.append(url) or _digest_response(),
    )

    client._get_form_digest()
    assert urls == [f"{get_config().sharepoint.site_url}/_api/contextinfo"]
    assert urls[0].endswith("/sites/VF_AIDV/_api/contextinfo")


def test_missing_digest_is_an_error_not_a_blank_header(monkeypatch):
    client = SharePointClient()
    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": "c"})
    monkeypatch.setattr("sharepoint.client.request_json", lambda url, **kw: {})
    with pytest.raises(Mcp365Error, match="FormDigest"):
        client._get_form_digest("https://t.sharepoint.com/sites/Eng")


def test_writes_go_to_the_drive_own_site_with_its_digest(monkeypatch):
    site = "https://t.sharepoint.com/sites/Eng"
    client = SharePointClient()
    calls = []

    def fake_request_json(url, headers=None, method="GET", data=None, context=""):
        calls.append((method, url, dict(headers or {})))
        if url.endswith("/_api/v2.0/drive"):
            return {"id": "drv1", "webUrl": f"{site}/Shared Documents"}
        if url.endswith("/_api/contextinfo"):
            return _digest_response()
        return {"id": "it1"}

    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": f"for:{host}"})
    monkeypatch.setattr(client, "get_token", lambda *a, **k: pytest.fail("the cookie channel works"))
    monkeypatch.setattr("sharepoint.client.request_json", fake_request_json)

    _info, drive_id = client.resolve_drive(f"{site}/Shared%20Documents/A")
    client.call_sharepoint_or_graph(f"/drives/{drive_id}/items/it1/content", method="PUT", data=b"x")

    assert [u for _m, u, _h in calls if u.endswith("/_api/contextinfo")] == [f"{site}/_api/contextinfo"]
    method, url, headers = calls[-1]
    assert (method, url) == ("PUT", f"{site}/_api/v2.0/drives/drv1/items/it1/content")
    assert headers["X-RequestDigest"] == "site-digest"


@pytest.fixture
def uploader(monkeypatch, tmp_path):
    """A client whose drive ``drv1`` lives in /sites/Eng; records every request."""
    site = "https://t.sharepoint.com/sites/Eng"
    client = SharePointClient()
    client._drive_sites["drv1"] = site
    client._drive_cache["web:drv1"] = f"{site}/Shared Documents"
    monkeypatch.setattr(client, "resolve_drive", lambda url="": ({}, "drv1"))
    monkeypatch.setattr(client, "ensure_folder", lambda drive, folder: None)
    monkeypatch.setattr(client, "_get_form_digest", lambda site_url="": f"digest-of:{site_url}")
    monkeypatch.setattr(client, "_cookie_headers", lambda accept="", host="": {"Cookie": f"for:{host}"})
    local = tmp_path / "bao cao's.txt"
    local.write_bytes(b"hello")
    return client, local


def test_upload_uses_rest_files_add_on_the_library_site(uploader, monkeypatch):
    client, local = uploader
    calls = []

    def fake_request_json(url, headers=None, method="GET", data=None, context=""):
        calls.append({"url": url, "headers": headers, "method": method, "data": data})
        return {
            "d": {
                "Name": local.name,
                "Length": "5",
                "ServerRelativeUrl": f"/sites/Eng/Shared Documents/S5/{local.name}",
                "UniqueId": "guid-1",
            }
        }

    monkeypatch.setattr(client, "get_token", lambda *a, **k: pytest.fail("the cookie upload works"))
    monkeypatch.setattr("sharepoint.client.request_json", fake_request_json)

    res = client.upload_file(str(local), "S5")

    assert len(calls) == 1
    call = calls[0]
    assert call["method"] == "POST"
    assert call["url"] == (
        "https://t.sharepoint.com/sites/Eng/_api/web/GetFolderByServerRelativeUrl"
        "('/sites/Eng/Shared%20Documents/S5')/Files/add(url='bao%20cao%27%27s.txt',overwrite=true)"
    )
    assert call["headers"]["X-RequestDigest"] == "digest-of:https://t.sharepoint.com/sites/Eng"
    assert call["headers"]["Cookie"] == "for:t.sharepoint.com"
    assert call["headers"]["Content-Type"] == "application/octet-stream"
    assert call["data"] == b"hello"
    assert res["size"] == 5 and res["id"] == "guid-1"
    assert res["webUrl"] == "https://t.sharepoint.com/sites/Eng/Shared%20Documents/S5/bao%20cao%27s.txt"


def test_upload_falls_back_to_graph_when_the_cookie_upload_fails(uploader, monkeypatch):
    client, local = uploader
    graph_calls = []

    def fake_request_json(url, headers=None, method="GET", data=None, context=""):
        if url.startswith("https://graph.microsoft.com"):
            graph_calls.append((method, url))
            return {"name": local.name, "size": 5, "id": "g1", "webUrl": "https://g"}
        raise AuthExpiredError("Không được xác thực (HTTP 401).")

    monkeypatch.setattr(client, "get_token", lambda *a, **k: "tok")
    monkeypatch.setattr("sharepoint.client.request_json", fake_request_json)

    assert client.upload_file(str(local), "S5")["id"] == "g1"
    assert graph_calls == [("PUT", "https://graph.microsoft.com/v1.0/drives/drv1/root:/S5/bao%20cao%27s.txt:/content")]



# ------------------------------------------------------------- delete


def _folder_target(client, monkeypatch):
    item = {
        "id": "it1",
        "name": "MHU full data",
        "size": 44786647,
        "folder": {"childCount": 3},
        "parentReference": {"path": "/drives/drv1/root:/S5/02. Tech's"},
    }
    monkeypatch.setattr(client, "resolve_file", lambda url: ("drv1", item))
    return client.describe_item("https://t.sharepoint.com/sites/Eng/Shared%20Documents/S5/x")


def test_describe_item_builds_the_server_relative_path_from_the_parent(uploader, monkeypatch):
    client, _ = uploader
    target = _folder_target(client, monkeypatch)
    assert target["path"] == "/sites/Eng/Shared Documents/S5/02. Tech's/MHU full data"
    assert target["is_folder"] and target["child_count"] == 3 and target["size"] == 44786647


def test_delete_recycles_by_default(uploader, monkeypatch):
    client, _ = uploader
    target = _folder_target(client, monkeypatch)
    calls = []
    monkeypatch.setattr(
        "sharepoint.client.request_json",
        lambda url, headers=None, method="GET", data=None, context="": calls.append((method, url, headers)) or {},
    )

    assert client.delete_item(target)["permanent"] is False
    method, url, headers = calls[0]
    assert method == "POST"
    assert url == (
        "https://t.sharepoint.com/sites/Eng/_api/web/GetFolderByServerRelativeUrl"
        "('/sites/Eng/Shared%20Documents/S5/02.%20Tech%27%27s/MHU%20full%20data')/recycle()"
    )
    assert headers["X-RequestDigest"] == "digest-of:https://t.sharepoint.com/sites/Eng"
    assert "X-HTTP-Method" not in headers


def test_permanent_delete_uses_delete_object_on_a_file(uploader, monkeypatch):
    client, _ = uploader
    target = {**_folder_target(client, monkeypatch), "is_folder": False, "path": "/sites/Eng/Shared Documents/a.png"}
    calls = []
    monkeypatch.setattr(
        "sharepoint.client.request_json",
        lambda url, headers=None, method="GET", data=None, context="": calls.append((method, url, headers)) or {},
    )

    client.delete_item(target, permanent=True)
    method, url, headers = calls[0]
    assert method == "POST"
    assert url.endswith("/GetFileByServerRelativeUrl('/sites/Eng/Shared%20Documents/a.png')")
    assert headers["X-HTTP-Method"] == "DELETE" and headers["IF-MATCH"] == "*"

# ------------------------------------------------------------- folders


@pytest.fixture
def folders(monkeypatch):
    """A drive holding the ``existing`` folders; records every GET and POST."""
    client = SharePointClient()
    existing: set[str] = set()
    calls: list[tuple[str, str, dict | None]] = []

    def fake_call_graph(path, method="GET", body=None, context=""):
        calls.append((method, path, body))
        if method == "POST":
            return {"id": "new", "folder": {}}
        rel = urllib.parse.unquote(path.split("root:/", 1)[1])
        if rel in existing:
            return {"id": rel, "folder": {}}
        raise Mcp365Error(f"HTTP 404 Not Found khi {context}.")

    monkeypatch.setattr(client, "call_graph", fake_call_graph)
    return client, existing, calls


def test_ensure_folder_does_not_post_when_the_folder_exists(folders):
    """A blind POST at the library root answered 403 although the folder was there."""
    client, existing, calls = folders
    existing.update({"A", "A/B"})
    client.ensure_folder("drv", "/sites/VF_AIDV/Shared Documents/A/B")
    assert [m for m, _p, _b in calls] == ["GET"]


def test_ensure_folder_creates_only_the_missing_part(folders):
    client, existing, calls = folders
    existing.add("A")
    client.ensure_folder("drv", "A/B/C")
    assert [(p, b["name"]) for m, p, b in calls if m == "POST"] == [
        ("https://graph.microsoft.com/v1.0/drives/drv/root:/A:/children", "B"),
        ("https://graph.microsoft.com/v1.0/drives/drv/root:/A/B:/children", "C"),
    ]
    # Nothing is looked up below a folder already known to be missing.
    assert [p for m, p, _b in calls if m == "GET"] == [
        "/drives/drv/root:/A/B/C",
        "/drives/drv/root:/A",
        "/drives/drv/root:/A/B",
    ]


def test_ensure_folder_reports_a_failed_lookup_instead_of_guessing(folders, monkeypatch):
    client, _existing, calls = folders

    def denied(path, method="GET", body=None, context=""):
        calls.append((method, path, body))
        raise AuthExpiredError("Bị từ chối truy cập (HTTP 403).", "Kiểm tra quyền.")

    monkeypatch.setattr(client, "call_graph", denied)
    with pytest.raises(AuthExpiredError):
        client.ensure_folder("drv", "A/B")
    assert [m for m, _p, _b in calls] == ["GET"]
