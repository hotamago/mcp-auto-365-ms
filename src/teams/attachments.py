"""Wire formats for Teams file cards and inline images."""

from __future__ import annotations

import html
import json
import re
from pathlib import Path
from typing import Any, Literal

from PIL import Image

from common.errors import Mcp365Error
from common.http import request, request_json

AttachmentMode = Literal["native", "link"]
AMS_OBJECTS = "https://as-api.asm.skype.com/v1/objects"
IMAGE_LIMIT = 20 * 1024 * 1024
_FORMATS = {".gif": "GIF", ".png": "PNG", ".jpeg": "JPEG", ".jpg": "JPEG"}
_OBJECT_ID = re.compile(r"[A-Za-z0-9_-]+\Z")


def inline_size(path: Path, image_as_file: bool) -> tuple[int, int] | None:
    """Return the display dimensions if the file qualifies for AMS, else use a card."""
    fmt = _FORMATS.get(path.suffix.lower())
    if image_as_file or not fmt or path.stat().st_size > IMAGE_LIMIT:
        return None
    try:
        with Image.open(path, formats=[fmt]) as image:
            w, h = image.size
            image.verify()
    except (ValueError, OSError, SyntaxError, Image.DecompressionBombError):
        return None
    return (w, h) if h <= 250 else (max(1, round(w * 250 / h)), 250)


def _object_id(value: Any) -> str:
    if not isinstance(value, str) or not _OBJECT_ID.fullmatch(value):
        raise Mcp365Error("AMS không trả về mã ảnh hợp lệ.", "Đọc lại tin trong Teams trước khi thử lại.")
    return value


def _create_object(token: str, path: str, body: dict) -> str:
    headers = {"Authorization": f"skype_token {token}", "Accept": "application/json", "Content-Type": "application/json",
               "x-ms-client-version": "1415/1.0.0.2018090100", "x-ams-post-sharing-mode": "Inline"}
    result = request_json(f"{AMS_OBJECTS}/{path}", method="POST", headers=headers,
                          data=json.dumps(body).encode(), max_retries=0, context="tạo ảnh AMS")
    return _object_id(result.get("id"))


def upload_image(token: str, chat: str, path: Path, size: tuple[int, int]) -> dict[str, Any]:
    with path.open("rb") as stream:
        data = stream.read(IMAGE_LIMIT + 1)
    if len(data) > IMAGE_LIMIT:
        raise Mcp365Error("File ảnh đã thay đổi và vượt 20 MiB.", "Xem lại bản nháp để gửi dạng thẻ tệp.")
    oid = _create_object(token, "", {"type": "pish/image", "permissions": {chat: ["read"]},
                                      "sharingMode": "Inline", "filename": path.name})
    request(f"{AMS_OBJECTS}/{oid}/content/imgpsh", method="PUT", data=data, kind="transfer",
            headers={"Authorization": f"skype_token {token}", "Content-Type": "application/octet-stream",
                     "x-ms-client-version": "1415/1.0.0.2018090100"}, context="tải ảnh lên AMS")
    url = f"{AMS_OBJECTS}/{oid}/views/imgo"
    tag = (f'<p><img src="{url}" itemtype="http://schema.skype.com/AMSImage" itemscope="{path.suffix[1:].lower()}" '
           f'width="{size[0]}" height="{size[1]}" alt="{html.escape(path.name)}" id="x_{oid}" itemid="{oid}"></p>')
    return {"ams_id": oid, "image_html": tag, "name": path.name, "size": len(data), "webUrl": url, "location": "Teams AMS"}


def file_card(item: dict, drive_id: str, name: str, share_url: str | None, share_id: str | None,
              scope: str | None) -> dict:
    ids = item.get("sharepointIds") or {}
    guid, url, site = ids.get("listItemUniqueId"), item.get("webDavUrl") or item.get("webUrl"), ids.get("siteUrl")
    if not all((guid, url, site, item.get("id"))):
        raise Mcp365Error(f"Thiếu metadata SharePoint cho thẻ tệp '{name}'.", 'Thử attachment_mode="link".')
    ext = Path(name).suffix[1:].lower()
    return {"@type": "http://schema.skype.com/File", "version": 2, "id": guid, "itemid": guid,
            "fileName": name, "title": name, "fileType": ext, "type": ext, "baseUrl": site,
            "objectUrl": url, "state": "active", "permissionScope": scope,
            "fileChicletState": {"serviceName": "teams" if scope is None else "p2p", "state": "active"},
            "sharepointIds": {**ids, "driveId": drive_id},
            "fileInfo": {"itemId": item["id"], "fileUrl": url, "siteUrl": site, "serverRelativeUrl": "",
                         "shareUrl": share_url, "shareId": share_id}}
