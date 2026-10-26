"""Cookie-only, ETag-guarded DOCX comment round trips.

File ETags are read from SP.File (not drive-item/list-item ETags). A stable
metadata/download/metadata snapshot prevents applying edits to mismatched bytes.
The sole write is a WebDAV PUT to the direct file URL with an exact If-Match.
SharePoint REST $value PUT ignores If-Match on this tenant (verified live);
it must never be used for conditional updates. Files/add(overwrite=true),
Graph and unconditional retries are never used by this comment operation.
"""

from __future__ import annotations

import json
import urllib.parse
from dataclasses import dataclass
from typing import Any

from common import approval
from common.errors import ConcurrentEditError, Mcp365Error
from common.http import request, request_json
from sharepoint import docx_comments
from sharepoint.client import SharePointClient, _odata_literal, view_url


@dataclass(frozen=True)
class Snapshot:
    endpoint: str
    site: str
    url: str
    etag: str
    version: str
    content: bytes


def _metadata(client: SharePointClient, endpoint: str) -> dict[str, str]:
    result = client._cookie_request(
        request_json,
        endpoint + "?$select=ETag,UIVersionLabel,ServerRelativeUrl,Name,UniqueId",
        headers={"Cache-Control": "no-cache"},
        context="read DOCX file ETag and version",
    )
    info = result.get("d", result)
    etag = info.get("ETag")
    if not isinstance(etag, str) or not etag.strip() or etag.strip() == "*" or etag.startswith("W/"):
        raise Mcp365Error(
            "SharePoint did not return a strong file ETag. Nothing uploaded.",
            "Open the direct file URL in Word Online; re-run the dry run when a file ETag is available.",
        )
    return info


def _conflict() -> ConcurrentEditError:
    return ConcurrentEditError(
        "File changed since download, or is locked (409/412/423). No unconditional overwrite was attempted.",
        "Re-run add_docx_comments with is_user_confirm=false to download the latest version, then approve again.",
    )


def snapshot(client: SharePointClient, url_or_guid: str) -> Snapshot:
    """Download a coherent file snapshot using the existing cookie/download layer."""
    with client.cookie_only():
        drive_id, item = client.resolve_file(url_or_guid)
        if not item.get("name", "").lower().endswith(".docx") or "folder" in item:
            raise Mcp365Error("add_docx_comments accepts .docx files only.", "Pass a Word .docx file URL or GUID.")
        site = client._drive_sites.get(drive_id, "")
        if not site:
            raise Mcp365Error("Cannot determine the file's owning site.", "Pass the direct SharePoint/OneDrive file URL.")
        path = client._server_relative_url(drive_id, item)
        endpoint = f"{site}/_api/web/GetFileByServerRelativeUrl('{_odata_literal(path)}')"
        before = _metadata(client, endpoint)
        try:
            _, content, _ = client.download(
                request, endpoint + "/$value", headers={"If-Match": before["ETag"], "Cache-Control": "no-cache"},
                kind="transfer", context="download DOCX for comments",
            )
        except ConcurrentEditError as exc:
            raise _conflict() from exc
        after = _metadata(client, endpoint)
        if before["ETag"] != after["ETag"]:
            raise _conflict()
        host = urllib.parse.urlparse(site).netloc
        url = f"https://{host}{urllib.parse.quote(path, safe='/')}"
        return Snapshot(endpoint, site, view_url(url), before["ETag"], str(before.get("UIVersionLabel") or "unknown"), content)


def upload(client: SharePointClient, source: Snapshot, content: bytes) -> None:
    """Exactly one conditional content update. Never refresh the precondition."""
    if not source.etag or source.etag.strip() == "*" or source.etag.startswith("W/"):
        raise Mcp365Error("Refusing upload without a strong exact ETag.", "Download a new snapshot and re-run the preview.")
    try:
        # snapshot() constructs this URL from the actual library and item path,
        # not from a sharing/viewer link. WebDAV enforces If-Match; both POST
        # with X-HTTP-Method: PUT and actual PUT to REST /$value ignored it live.
        file_url = urllib.parse.urlparse(source.url)._replace(query="", fragment="").geturl()
        client._cookie_request(
            request, file_url, method="PUT",
            headers={"If-Match": source.etag, "Content-Type": "application/octet-stream"},
            digest_site=source.site, data=content, max_retries=0,
            kind="transfer", context="conditionally upload DOCX comments as a new version",
        )
    except ConcurrentEditError as exc:
        raise _conflict() from exc


def add_comments(
    client: SharePointClient,
    url_or_guid: str,
    comments: list[dict[str, Any]],
    is_user_confirm: bool,
    author: str = "",
) -> str:
    with client.cookie_only():
        source = snapshot(client, url_or_guid)
        if not author.strip():
            # SP.CurrentUser is derived from the same authenticated browser session
            # as the document. It works even when Teams and Graph are unavailable.
            identity = client._cookie_request(
                request_json, source.site + "/_api/web/currentuser?$select=Title",
                context="derive signed-in DOCX comment author",
            )
            author = identity.get("d", identity).get("Title", "").strip()
        if not author.strip():
            raise Mcp365Error("Cannot derive the signed-in user's display name.", "Provide author explicitly or renew the browser session.")
        edited = docx_comments.prepare(source.content, comments, author)
        preview = (
            f"File: {source.url}\nETag: {source.etag}\nVersion: {source.version}\n"
            f"Author: {author}\nInitials: {''.join(w[0] for w in author.split('(')[0].split()).upper()}\n"
            + json.dumps(edited.preview, ensure_ascii=False, indent=2)
        )
        # False is intentionally a read-only preview; reads are needed to resolve
        # anchors/threads. Literal true is still required for the sole mutation.
        if is_user_confirm is not True:
            return "CHƯA GỬI — dry run; nothing uploaded.\n" + preview + "\nShow this exact draft to the user and obtain approval before is_user_confirm=true."
        approval.require_confirm(is_user_confirm, "Add DOCX comments", source.url, preview)
        upload(client, source, edited.data)
        try:
            checked = snapshot(client, url_or_guid)
            docx_comments.verify(checked.content, edited.records)
            if checked.etag == source.etag:
                raise ValueError("file ETag did not change after upload")
        except Exception as exc:
            raise Mcp365Error(
                f"DOCX upload succeeded, but verification failed: {exc}",
                f"The file may already contain these comments. Open {source.url} and inspect before retrying; do not resend blindly.",
            ) from exc
        return (
            f"Added and re-downloaded/verified {len(edited.records)} comment(s).\n"
            f"File: {checked.url}\nVersion: {checked.version}\nETag: {checked.etag}\n"
            + json.dumps(edited.preview, ensure_ascii=False, indent=2)
        )
