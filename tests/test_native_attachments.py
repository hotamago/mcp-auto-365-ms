"""Native attachments use a different message wire format and storage ACL than link mode."""

import json
from urllib.parse import unquote

import pytest
from PIL import Image

from common.errors import ConnectError, Mcp365Error, TransportError
from sharepoint.client import SharePointClient
from teams import attachments
from teams.client import TeamsClient


@pytest.fixture
def client(monkeypatch, identity):
    client = TeamsClient()
    monkeypatch.setattr(type(client), "identity", property(lambda self: identity))
    monkeypatch.setattr(client, "_auth", lambda: {"token": "test"})
    return client


def posted(monkeypatch, client):
    posts = []
    monkeypatch.setattr(client, "_chat_json", lambda method, url, **kw: posts.append(json.loads(kw["data"])) or {})
    monkeypatch.setattr(client, "_resolve_sent_id", lambda *args: "99")
    return posts


def png(tmp_path):
    path = tmp_path / "photo.png"
    Image.new("RGB", (17, 3)).save(path)
    return path


def test_inline_create_upload_send(client, tmp_path, monkeypatch):
    path = png(tmp_path)
    calls = []
    monkeypatch.setattr(attachments, "request_json", lambda url, **kw: calls.append((url, kw)) or {"id": "image-a"})
    monkeypatch.setattr(attachments, "request", lambda url, **kw: calls.append((url, kw)))
    monkeypatch.setattr(client, "_chat_file_plan", lambda *a: pytest.fail("image-only never plans OneDrive sharing"))
    posts = posted(monkeypatch, client)
    client.send_message("48:notes", "photo", file_path=str(path))
    (create, post), (upload, put) = calls
    assert (create, post["method"]) == (attachments.AMS_OBJECTS + "/", "POST")
    assert json.loads(post["data"]) == {"type": "pish/image", "permissions": {"48:notes": ["read"]}, "sharingMode": "Inline", "filename": path.name}
    assert (upload, put["method"], put["data"]) == (attachments.AMS_OBJECTS + "/image-a/content/imgpsh", "PUT", path.read_bytes())
    assert post["headers"]["Authorization"] == put["headers"]["Authorization"] == "skype_token test"
    assert put["headers"]["Content-Type"] == "application/octet-stream" and post["headers"]["x-ams-post-sharing-mode"] == "Inline"
    assert posts[0]["amsreferences"] == ["image-a"] and "amsreferences" not in posts[0].get("properties", {})
    assert 'width="17" height="3"' in posts[0]["content"] and 'itemid="image-a"' in posts[0]["content"]


def test_format_limit_and_image_as_file(tmp_path):
    image = png(tmp_path)
    assert attachments.inline_size(image, True) is None
    with image.open("ab") as f:
        f.truncate(attachments.IMAGE_LIMIT + 1)
    assert attachments.inline_size(image, False) is None
    image.write_bytes(b"not a PNG")
    assert attachments.inline_size(image, False) is None


def test_native_file_metadata_and_permission_match(client, monkeypatch, tmp_path):
    file = tmp_path / "file.txt"
    file.write_text("file")
    site = "https://tenant-my.sharepoint.com/personal/me"
    share = site + "/share/exact"
    sp = SharePointClient()
    sp._drive_sites["drive"] = site
    meta = {"id": "item", "webDavUrl": site + "/Documents/file.txt", "sharepointIds": {"siteUrl": site, "listItemUniqueId": "uid"}}
    urls = []

    def read(fn, url, **kw):
        urls.append(url)
        if "/permissions" in url:
            return {"value": [{"id": "wrong", "link": {"webUrl": share + "2"}}, {"id": "right", "link": {"webUrl": share + "?email=me"}}]}
        return meta

    monkeypatch.setattr(sp, "_cookie_request", read)
    monkeypatch.setattr("sharepoint.client.SharePointClient", lambda: sp)
    monkeypatch.setattr(sp, "upload_chat_file", lambda p: {"id": "uid", "drive_id": "drive", "name": file.name, "webUrl": site})
    monkeypatch.setattr(sp, "create_people_link", lambda d, path, upns, access: share if upns == [client.identity.upn] else pytest.fail("wrong recipients"))
    posts = posted(monkeypatch, client)
    client.send_message("48:notes", "text", file_path=str(file))
    assert urls == [f"{site}/_api/v2.0/drives/drive/items/uid?$select=id,name,webUrl,webDavUrl,sharepointIds", f"{site}/_api/v2.0/drives/drive/items/uid/permissions"]
    assert posts[0]["content"] == "<p>text</p>"
    card, = json.loads(posts[0]["properties"]["files"])
    assert (card["id"], card["itemid"], card["version"], card["fileChicletState"]["serviceName"]) == ("uid", "uid", 2, "p2p")
    assert card["fileInfo"] == {"itemId": "item", "fileUrl": meta["webDavUrl"], "siteUrl": site, "serverRelativeUrl": "", "shareId": "right", "shareUrl": share}
    assert card["permissionScope"] == "SpecificPeople" and card["sharepointIds"]["driveId"] == "drive"


def test_channel_uses_verified_folder_no_sharing(client, monkeypatch, tmp_path):
    chat = "19:channel@thread.tacv2"
    site = "https://tenant.sharepoint.com/sites/S"
    target = []
    file = tmp_path / "f.txt"
    file.write_text("private")
    sp = SharePointClient()
    monkeypatch.setattr("sharepoint.client.SharePointClient", lambda: sp)
    p = {"siteId": "tenant.sharepoint.com,s,w", "driveId": "drive", "channelDocsFolderRelativeUrl": "/sites/S/Custom library/Channel"}
    monkeypatch.setattr(client, "_chat_json", lambda m, url, **kw: {"properties": p} if m == "GET" and unquote(url).startswith("/threads/" + chat) else pytest.fail("expected thread read"))
    monkeypatch.setattr(sp, "resolve_drive", lambda url: ({"site_path": "/sites/S"}, "drive"))
    monkeypatch.setattr(sp, "_drive_web_url", lambda d: site + "/Custom%20library")
    monkeypatch.setattr(sp, "upload_file", lambda path, target_folder_url_or_path: target.append(target_folder_url_or_path) or {"webUrl": site, "drive_id": "drive", "name": file.name})
    meta = {"id": "i", "webUrl": site, "sharepointIds": {"siteUrl": site, "listItemUniqueId": "uid"}}
    monkeypatch.setattr(sp, "card_metadata", lambda info, share: (meta, None) if share is None else pytest.fail("channel share"))
    info, = client._attach_files({"id": chat}, [file], "members", False, "native", False)
    assert target == [site + "?id=Channel"] and info["card"]["fileChicletState"]["serviceName"] == "teams"
    assert info["card"]["permissionScope"] is info["card"]["fileInfo"]["shareId"] is info["card"]["fileInfo"]["shareUrl"] is None
    p.pop("driveId")
    with pytest.raises(Mcp365Error, match="attachment_mode"):
        client._attach_files({"id": chat}, [file], "members", False, "native", False)
    assert len(target) == 1


def test_failed_image_discards_completed_file_not_ams(client, monkeypatch, tmp_path):
    file = tmp_path / "f.txt"
    file.write_text("txt")
    path = png(tmp_path)
    discards = []
    monkeypatch.setattr(client, "_upload_chat_file", lambda *a: {"name": file.name, "location": "OneDrive", "share_link": "u", "drive_id": "d"})
    monkeypatch.setattr(SharePointClient, "card_metadata", lambda *a: ({"id": "i", "webUrl": "u", "sharepointIds": {"listItemUniqueId": "uid", "siteUrl": "s"}}, "perm"))
    monkeypatch.setattr(attachments, "upload_image", lambda *a: (_ for _ in ()).throw(Mcp365Error("AMS rejected")))
    monkeypatch.setattr(SharePointClient, "discard_uploaded", lambda self, items: discards.extend(i["name"] for i in items) or (discards, []))
    posts = posted(monkeypatch, client)
    with pytest.raises(Mcp365Error, match="1/2"):
        client.send_message("48:notes", "text", file_paths=[str(path), str(file)])
    assert discards == [file.name] and not posts


@pytest.mark.parametrize("error,status,deleted", [(ConnectError, None, True), (Mcp365Error, 403, True), (Mcp365Error, 503, False), (TransportError, None, False)])
def test_cleanup_depends_on_delivery_certainty(monkeypatch, error, status, deleted):
    discarded = []
    monkeypatch.setattr(SharePointClient, "discard_uploaded", lambda self, files: discarded.extend(f["name"] for f in files) or (discarded, []))
    exc = error("not sent or lost response")
    exc.http_status = status
    TeamsClient._unsent_with_files(exc, [{"name": "txt", "location": "OneDrive"}, {"name": "png", "location": "Teams AMS"}])
    assert discarded == (["txt"] if deleted else [])
