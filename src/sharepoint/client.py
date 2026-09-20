"""SharePoint & OneDrive client.

Two independent channels, because neither one alone covers the job:

**Graph (Azure CLI token)** - used for writes (folder creation, upload,
replace) and for drive/item metadata.

**Direct session cookies (``rtFa`` + ``FedAuth``)** - used for binary
downloads, the REST search API and version history, which Graph either blocks
or does not expose to a developer token.
"""

from __future__ import annotations

import difflib
import io
import json
import re
import shutil
import subprocess
import tempfile
import time
import urllib.parse
import zipfile
from pathlib import Path
from typing import Any

from common.chrome_cookies import ChromeCookieDecryptor
from common.config import get_config
from common.errors import (
    AuthExpiredError,
    CAEChallengeError,
    Mcp365Error,
    UnsupportedOperationError,
)
from common.http import request, request_bytes, request_json

_GUID_RE = re.compile(r"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})")
_SHARED_DOCS_RE = re.compile(r"^/sites/[^/]+/Shared Documents/?", re.IGNORECASE)

_BINARY_EXTS = (".docx", ".xlsx", ".pptx", ".pdf", ".7z", ".zip", ".txt", ".md", ".csv", ".json", ".xml", ".png", ".jpg")
_TEXT_EXTS = (".txt", ".md", ".csv", ".json", ".xml", ".py", ".yaml", ".yml", ".ts", ".js", ".java", ".c", ".h", ".sql")
_WORD_EXTS = (".docx", ".dotx")
_SHEET_EXTS = (".xlsx", ".xlsm")


def _strip_library_prefix(path: str) -> str:
    return _SHARED_DOCS_RE.sub("", path or "").strip("/")


def human_size(size: int) -> str:
    if size > 1024 * 1024:
        return f"{size / (1024 * 1024):.2f} MB"
    if size > 1024:
        return f"{size / 1024:.1f} KB"
    return f"{size} B"


class SharePointClient:
    def __init__(self) -> None:
        self._token: str | None = None
        self._token_expires: float = 0.0
        self._site_cache: dict[str, str] = {}
        self._drive_cache: dict[str, str] = {}

    # ------------------------------------------------------------ Graph auth

    def get_token(self, force_refresh: bool = False) -> str:
        now = time.time()
        if not force_refresh and self._token and now < (self._token_expires - 60):
            return self._token
        try:
            res = subprocess.run(
                ["az", "account", "get-access-token", "--resource", "https://graph.microsoft.com", "-o", "json"],
                capture_output=True,
                text=True,
                check=True,
                timeout=60,
            )
        except FileNotFoundError as exc:
            raise UnsupportedOperationError(
                "Không tìm thấy Azure CLI ('az') trên máy.",
                "Cài Azure CLI rồi chạy `az login`. Các tool chỉ đọc/tải file vẫn hoạt động nhờ cookie Chrome.",
            ) from exc
        except subprocess.TimeoutExpired as exc:
            raise Mcp365Error("Azure CLI không phản hồi sau 60s.", "Thử chạy `az account get-access-token` thủ công.") from exc
        except subprocess.CalledProcessError as exc:
            raise AuthExpiredError(
                f"Azure CLI không cấp được token: {(exc.stderr or '').strip()[:300]}",
                "Chạy: az login --scope https://graph.microsoft.com/.default",
            ) from exc

        data = json.loads(res.stdout)
        self._token = data["accessToken"]
        expires_on = data.get("expires_on")
        self._token_expires = float(expires_on) if isinstance(expires_on, int | float) else now + 3000
        return self._token

    def call_graph(self, path: str, method: str = "GET", body: dict | None = None, context: str = "") -> dict:
        url = path if path.startswith("http") else f"https://graph.microsoft.com/v1.0{path}"
        payload = json.dumps(body).encode("utf-8") if body is not None else None

        def headers() -> dict[str, str]:
            hdrs = {"Authorization": f"Bearer {self.get_token()}", "Accept": "application/json"}
            if payload is not None:
                hdrs["Content-Type"] = "application/json"
            return hdrs

        try:
            return request_json(url, headers=headers(), method=method, data=payload, context=context or f"gọi Graph {path}")
        except CAEChallengeError:
            # Retrying is pointless: the Azure CLI hands back the byte-identical
            # cached token until it truly expires, so a CAE challenge can only be
            # cleared by an interactive login.
            raise
        except AuthExpiredError:
            self._token = None
            self._token_expires = 0.0
            self.get_token(force_refresh=True)
            return request_json(url, headers=headers(), method=method, data=payload, context=context or f"gọi Graph {path}")

    # --------------------------------------------------------- cookie auth

    def _cookie_headers(self, accept: str = "application/json;odata=verbose") -> dict[str, str]:
        """Headers for the direct-session channel.

        Cookie order (``rtFa`` first) and a browser User-Agent are both required;
        the previous code applied them only on the download path, so search and
        version history went out malformed.
        """
        cookies = ChromeCookieDecryptor.get_cookies_for_domain("sharepoint.com", ["rtFa", "FedAuth"])
        if not cookies.get("FedAuth") or not cookies.get("rtFa"):
            raise AuthExpiredError(
                "Không tìm thấy cookie phiên SharePoint (rtFa/FedAuth) trong Chrome.",
                "Mở https://<tenant>.sharepoint.com trong Chrome, đăng nhập và tick 'Stay signed in'.",
            )
        return {
            "Cookie": f"rtFa={cookies['rtFa']}; FedAuth={cookies['FedAuth']}",
            "User-Agent": get_config().http.user_agent,
            "Accept": accept,
        }

    # ------------------------------------------------------------- resolving

    def parse_sharepoint_url(self, url: str) -> dict[str, Any]:
        cfg = get_config().sharepoint
        parsed = urllib.parse.urlparse(url)
        path = urllib.parse.unquote(parsed.path)
        query = urllib.parse.parse_qs(parsed.query)

        info: dict[str, Any] = {
            "hostname": parsed.netloc or cfg.hostname,
            "site_name": None,
            "site_path": None,
            "folder_path": None,
            "sourcedoc": None,
            "file_name": query.get("file", [None])[0],
            "is_personal": "-my.sharepoint.com" in (parsed.netloc or ""),
            "type": "unknown",
        }

        if "sourcedoc" in query:
            info["sourcedoc"] = re.sub(r"[{}]", "", query["sourcedoc"][0])
            info["type"] = "document"

        site_match = re.search(r"/sites/([^/]+)", path)
        if site_match:
            info["site_name"] = site_match.group(1)
            info["site_path"] = f"/sites/{site_match.group(1)}"
        else:
            info["site_name"] = cfg.site_name
            info["site_path"] = cfg.site_path

        if "id" in query:
            info["folder_path"] = query["id"][0]
            if not info["sourcedoc"]:
                info["type"] = "folder"

        if info["type"] == "unknown":
            if any(m in path for m in (":w:", ":x:", ":p:", ":b:")) or path.lower().endswith(_BINARY_EXTS):
                info["type"] = "document"
        return info

    def get_site_id(self, hostname: str, site_path: str) -> str:
        key = f"{hostname}:{site_path}"
        if key not in self._site_cache:
            data = self.call_graph(f"/sites/{hostname}:{site_path}", context=f"tra cứu site {site_path}")
            self._site_cache[key] = data["id"]
        return self._site_cache[key]

    def get_default_drive_id(self, site_id: str) -> str:
        if site_id not in self._drive_cache:
            drives = self.call_graph(f"/sites/{site_id}/drives", context="liệt kê thư viện tài liệu").get("value", [])
            if not drives:
                raise Mcp365Error(f"Site {site_id} không có thư viện tài liệu nào.", "Kiểm tra lại site đích.")
            self._drive_cache[site_id] = drives[0]["id"]
        return self._drive_cache[site_id]

    def resolve_drive(self, url: str = "") -> tuple[dict[str, Any], str]:
        """Return ``(url_info, drive_id)`` for a URL, falling back to config."""
        info = self.parse_sharepoint_url(url) if url.startswith("http") else self._default_info()
        site_id = self.get_site_id(info["hostname"], info["site_path"])
        return info, self.get_default_drive_id(site_id)

    def _default_info(self) -> dict[str, Any]:
        cfg = get_config().sharepoint
        return {
            "hostname": cfg.hostname,
            "site_name": cfg.site_name,
            "site_path": cfg.site_path,
            "folder_path": None,
            "sourcedoc": None,
            "file_name": None,
            "is_personal": False,
            "type": "unknown",
        }

    # ------------------------------------------------------------- browsing

    def list_folder_contents(self, drive_id: str, relative_path: str) -> list[dict[str, Any]]:
        clean = _strip_library_prefix(relative_path)
        endpoint = (
            f"/drives/{drive_id}/root:/{urllib.parse.quote(clean)}:/children" if clean else f"/drives/{drive_id}/root/children"
        )
        return self.call_graph(endpoint, context=f"liệt kê thư mục '{clean or '/'}'").get("value", [])

    def get_item_by_guid(self, drive_id: str, guid: str) -> dict[str, Any]:
        return self.call_graph(f"/drives/{drive_id}/items/{guid}", context=f"đọc metadata item {guid}")

    def get_item_versions(self, drive_id: str, item_id: str) -> list[dict[str, Any]]:
        return self.call_graph(f"/drives/{drive_id}/items/{item_id}/versions", context="đọc lịch sử phiên bản").get(
            "value", []
        )

    def read_link(self, url: str, max_depth: int = 2) -> str:
        info, drive_id = self.resolve_drive(url)
        if info["type"] == "document" and info["sourcedoc"]:
            return self.describe_document(drive_id, info["sourcedoc"], info.get("file_name"))
        return self._render_folder_tree(drive_id, info.get("folder_path") or "", max_depth=max_depth)

    def describe_document(self, drive_id: str, guid: str, file_name: str | None = None) -> str:
        item = self.get_item_by_guid(drive_id, guid)
        size = item.get("size", 0)
        parent = item.get("parentReference", {}).get("path", "")
        if "root:" in parent:
            parent = parent.split("root:")[-1]

        out = [
            f"# Document: {item.get('name', file_name or 'Unknown')}",
            f"- **Path in SharePoint**: `{parent}/{item.get('name', '')}`",
            f"- **GUID / UniqueId**: `{guid}`",
            f"- **Size**: {size:,} bytes ({human_size(size)})",
            f"- **Created By**: {item.get('createdBy', {}).get('user', {}).get('displayName', 'Unknown')} "
            f"({item.get('createdDateTime', '')})",
            f"- **Last Modified By**: {item.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'Unknown')} "
            f"({item.get('lastModifiedDateTime', '')})",
            f"- **Direct Web URL**: {item.get('webUrl')}",
            "",
            "## Version History",
        ]
        try:
            versions = self.get_item_versions(drive_id, guid)
        except Mcp365Error as exc:
            out.append(f"*Không đọc được lịch sử phiên bản: {exc.message}*")
            return "\n".join(out)

        if not versions:
            out.append("*No version history available*")
        else:
            out.append("| Version | Modified Time | Modified By | Size |")
            out.append("| --- | --- | --- | --- |")
            for v in versions[:20]:
                out.append(
                    f"| {v.get('id')} | {v.get('lastModifiedDateTime', '')[:19].replace('T', ' ')} "
                    f"| {v.get('lastModifiedBy', {}).get('user', {}).get('displayName', 'Unknown')} "
                    f"| {v.get('size', 0):,} B |"
                )
            if len(versions) > 20:
                out.append(f"| ... | and {len(versions) - 20} older versions | | |")

        out.append(f"\n> **Download**: Call `download_sharepoint_link('{guid}')` to download the original file.")
        return "\n".join(out)

    def _render_folder_tree(self, drive_id: str, folder_path: str, max_depth: int = 2) -> str:
        out = [f"# SharePoint Folder: `{folder_path or '/'}`\n"]

        def traverse(rel_path: str, depth: int, prefix: str = "") -> None:
            if depth > max_depth:
                return
            try:
                items = self.list_folder_contents(drive_id, rel_path)
            except Mcp365Error as exc:
                out.append(f"{prefix}- *[Lỗi khi tải: {exc.message}]*")
                return

            folders = sorted((i for i in items if "folder" in i), key=lambda x: x["name"].lower())
            files = sorted((i for i in items if "folder" not in i), key=lambda x: x["name"].lower())
            for f in folders:
                out.append(f"{prefix}📁 **{f['name']}/** *({f.get('folder', {}).get('childCount', 0)} items)*")
                traverse(f"{rel_path}/{f['name']}".strip("/"), depth + 1, prefix + "  ")
            for doc in files:
                user = doc.get("lastModifiedBy", {}).get("user", {}).get("displayName", "")
                out.append(
                    f"{prefix}📄 `{doc['name']}` *({human_size(doc.get('size', 0))}, "
                    f"{doc.get('lastModifiedDateTime', '')[:10]}{f', by {user}' if user else ''})*"
                )

        traverse(folder_path, 1)
        out.append("\n> **Download All**: Call `download_sharepoint_link(url)` to download all documents in this folder.")
        return "\n".join(out)

    # ------------------------------------------------------------ downloads

    def _download_to(self, file_url: str, dest: Path, headers: dict[str, str], results: list[tuple[str, int, str]]) -> bool:
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            data = request_bytes(file_url, headers=headers, context=f"tải '{dest.name}'")
        except Mcp365Error as exc:
            results.append((dest.name, 0, f"Lỗi: {exc.message}"))
            return False
        dest.write_bytes(data)
        results.append((dest.name, len(data), str(dest)))
        return True

    def download_link(self, url_or_guid: str, target_dir: str = "") -> str:
        cfg = get_config().sharepoint
        target_path = Path(target_dir or cfg.default_download_dir)
        target_path.mkdir(parents=True, exist_ok=True)
        headers = {**self._cookie_headers(accept="*/*")}
        results: list[tuple[str, int, str]] = []

        if url_or_guid.startswith("http"):
            self._download_from_url(url_or_guid, target_path, headers, results)
        elif _GUID_RE.search(url_or_guid):
            guid = _GUID_RE.search(url_or_guid).group(1)
            info, drive_id = self.resolve_drive()
            item = self.get_item_by_guid(drive_id, guid)
            parent = item.get("parentReference", {}).get("path", "").split("root:")[-1]
            server_rel = f"{info['site_path']}/Shared Documents{parent}/{item['name']}"
            self._download_to(
                f"https://{info['hostname']}{urllib.parse.quote(server_rel)}",
                target_path / item["name"],
                headers,
                results,
            )
        else:
            raise Mcp365Error(
                f"Không nhận dạng được link hoặc GUID SharePoint: {url_or_guid}",
                "Truyền URL đầy đủ (https://...) hoặc UniqueId dạng GUID.",
            )
        return self._render_download_report(results, str(target_path))

    def _download_from_url(
        self, url: str, target_path: Path, headers: dict[str, str], results: list[tuple[str, int, str]]
    ) -> None:
        # 1. Personal OneDrive sharing link -> resolve via the WOPI context.
        if "-my.sharepoint.com" in url and any(m in url for m in (":x:/", ":w:/", ":p:/", ":b:/")):
            page = request_bytes(url, headers=headers, context="mở link chia sẻ OneDrive").decode("utf-8", errors="ignore")
            match = re.search(r"var _wopiContextJson\s*=\s*({.*?});", page)
            if not match:
                raise Mcp365Error(
                    "Không trích xuất được thông tin file từ link chia sẻ OneDrive.",
                    "Link có thể đã hết hạn, hoặc cần mở bằng trình duyệt trước.",
                )
            wopi = json.loads(match.group(1))
            file_url = wopi.get("FileGetUrl")
            if not file_url:
                raise Mcp365Error("Link chia sẻ không chứa FileGetUrl.", "Thử tải trực tiếp bằng đường dẫn đầy đủ của file.")
            self._download_to(file_url, target_path / (wopi.get("FileName") or "downloaded_file"), headers, results)
            return

        parsed_path = urllib.parse.unquote(urllib.parse.urlparse(url).path)

        # 2. Direct path to a file -> fetch bytes straight from the web server.
        #    NOTE: this branch previously fell through into the Graph folder
        #    logic because `drive_id = ...` was indented one level too far out,
        #    raising UnboundLocalError *after* the file had been written.
        if parsed_path.lower().endswith(_BINARY_EXTS) and "/sites/" in parsed_path:
            hostname = urllib.parse.urlparse(url).netloc
            file_url = f"https://{hostname}{urllib.parse.quote(parsed_path, safe='/:')}"
            self._download_to(file_url, target_path / parsed_path.split("/")[-1], headers, results)
            return

        # 3. Anything else -> resolve through Graph (single document or folder).
        info, drive_id = self.resolve_drive(url)
        if info["type"] == "document" and info["sourcedoc"]:
            item = self.get_item_by_guid(drive_id, info["sourcedoc"])
            parent = item.get("parentReference", {}).get("path", "").split("root:")[-1]
            server_rel = f"{info['site_path']}/Shared Documents{parent}/{item['name']}"
            self._download_to(
                f"https://{info['hostname']}{urllib.parse.quote(server_rel)}",
                target_path / item["name"],
                headers,
                results,
            )
            return

        clean = _strip_library_prefix(info.get("folder_path") or "")
        endpoint = f"/drives/{drive_id}/root:/{urllib.parse.quote(clean)}" if clean else f"/drives/{drive_id}/root"
        folder_item = self.call_graph(endpoint, context="mở thư mục SharePoint")
        self._sync_folder_down(
            drive_id,
            folder_item["id"],
            target_path,
            f"{info['site_path']}/Shared Documents/{clean}".rstrip("/"),
            info["hostname"],
            headers,
            results,
        )

    def _sync_folder_down(
        self,
        drive_id: str,
        folder_id: str,
        local_dir: Path,
        server_parent: str,
        hostname: str,
        headers: dict[str, str],
        results: list[tuple[str, int, str]],
        skip_large_media: bool = True,
    ) -> None:
        items = self.call_graph(f"/drives/{drive_id}/items/{folder_id}/children", context="liệt kê nội dung thư mục").get(
            "value", []
        )
        for item in items:
            name = item["name"]
            if "folder" in item:
                self._sync_folder_down(
                    drive_id, item["id"], local_dir / name, f"{server_parent}/{name}", hostname, headers, results
                )
                continue
            if skip_large_media and name.lower().endswith((".mp4", ".mov")) and item.get("size", 0) > 50 * 1024 * 1024:
                results.append((name, 0, "Bỏ qua: video lớn hơn 50MB"))
                continue
            file_url = f"https://{hostname}{urllib.parse.quote(f'{server_parent}/{name}')}"
            self._download_to(file_url, local_dir / name, headers, results)

    def _render_download_report(self, results: list[tuple[str, int, str]], target_dir: str) -> str:
        ok = [r for r in results if r[1] > 0]
        out = [f"# Downloaded {len(ok)}/{len(results)} files to `{target_dir}`\n"]
        out.append("| File Name | Size | Local Path |")
        out.append("| --- | --- | --- |")
        for name, size, path in results:
            out.append(f"| `{name}` | {human_size(size)} | `{path}` |" if size else f"| `{name}` | — | {path} |")
        return "\n".join(out)

    def download_meeting_recordings(self, target_dir: str = "", limit: int = 5, query: str = "Recording") -> str:
        """Download Teams meeting recordings, which land in OneDrive/SharePoint.

        Uses the cookie channel, so it needs no Graph scope.
        """
        hits = self.search_files(query=query, max_results=max(limit * 3, 10))
        videos = [h for h in hits if h["path"].lower().endswith((".mp4", ".mov"))][:limit]
        if not videos:
            return f"Không tìm thấy bản ghi cuộc họp nào khớp với '{query}'."

        cfg = get_config().sharepoint
        target_path = Path(target_dir or cfg.default_download_dir) / "recordings"
        target_path.mkdir(parents=True, exist_ok=True)
        headers = self._cookie_headers(accept="*/*")
        results: list[tuple[str, int, str]] = []
        for video in videos:
            parsed = urllib.parse.urlparse(video["path"])
            file_url = f"https://{parsed.netloc}{urllib.parse.quote(urllib.parse.unquote(parsed.path), safe='/:')}"
            self._download_to(file_url, target_path / video["title"], headers, results)
        return self._render_download_report(results, str(target_path))

    # -------------------------------------------------------------- uploads

    def ensure_folder(self, drive_id: str, folder_path: str) -> None:
        parts = [p for p in _strip_library_prefix(folder_path).split("/") if p]
        current = ""
        for part in parts:
            parent = f"root:/{urllib.parse.quote(current, safe='/')}:" if current else "root"
            try:
                self.call_graph(
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/{parent}/children",
                    method="POST",
                    body={"name": part, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
                    context=f"tạo thư mục '{part}'",
                )
            except Mcp365Error as exc:
                # Already existing is the expected, benign outcome.
                if "nameAlreadyExists" not in str(exc) and "409" not in str(exc):
                    raise
            current = f"{current}/{part}" if current else part

    def upload_file(self, local_file_path: str, target_folder_url_or_path: str, target_file_name: str | None = None) -> dict:
        local = Path(local_file_path).expanduser().resolve()
        if not local.is_file():
            raise Mcp365Error(f"Không tìm thấy file cần upload: {local_file_path}", "Kiểm tra lại đường dẫn.")

        file_name = target_file_name or local.name
        size = local.stat().st_size

        if target_folder_url_or_path.startswith("http"):
            info, drive_id = self.resolve_drive(target_folder_url_or_path)
            folder_path = info.get("folder_path") or ""
        else:
            _info, drive_id = self.resolve_drive()
            folder_path = target_folder_url_or_path

        clean_folder = _strip_library_prefix(folder_path)
        if clean_folder:
            self.ensure_folder(drive_id, clean_folder)
        remote_path = f"{clean_folder}/{file_name}" if clean_folder else file_name
        encoded = urllib.parse.quote(remote_path, safe="/")

        if size <= 100 * 1024 * 1024:
            data = request_json(
                f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded}:/content",
                headers={"Authorization": f"Bearer {self.get_token()}", "Content-Type": "application/octet-stream"},
                method="PUT",
                data=local.read_bytes(),
                context=f"upload '{file_name}'",
            )
        else:
            data = self._upload_large(drive_id, encoded, local, size)

        return {
            "status": "UPLOADED",
            "name": data.get("name", file_name),
            "size": data.get("size", size),
            "id": data.get("id"),
            "webUrl": data.get("webUrl"),
            "folder": clean_folder or "/",
        }

    def _upload_large(self, drive_id: str, encoded_path: str, local: Path, size: int) -> dict:
        session = self.call_graph(
            f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}:/createUploadSession",
            method="POST",
            body={"item": {"@microsoft.graph.conflictBehavior": "replace"}},
            context="tạo phiên upload lớn",
        )
        upload_url = session["uploadUrl"]
        chunk_size = 10 * 1024 * 1024
        sent = 0
        result: dict = {}
        with local.open("rb") as fh:
            while sent < size:
                blob = fh.read(chunk_size)
                if not blob:
                    break
                status, body, _hdrs = request(
                    upload_url,
                    headers={
                        "Content-Length": str(len(blob)),
                        "Content-Range": f"bytes {sent}-{sent + len(blob) - 1}/{size}",
                    },
                    method="PUT",
                    data=blob,
                    context=f"upload chunk {sent // chunk_size + 1}",
                )
                if status in (200, 201) and body:
                    result = json.loads(body.decode("utf-8"))
                sent += len(blob)
        return result

    def replace_file(self, local_file_path: str, file_url_or_guid: str) -> dict:
        local = Path(local_file_path).expanduser().resolve()
        if not local.is_file():
            raise Mcp365Error(f"Không tìm thấy file: {local_file_path}", "Kiểm tra lại đường dẫn.")

        if file_url_or_guid.startswith("http"):
            info, drive_id = self.resolve_drive(file_url_or_guid)
            if info.get("sourcedoc"):
                endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{info['sourcedoc']}/content"
            else:
                clean = _strip_library_prefix(info.get("folder_path") or info.get("file_name") or "")
                endpoint = (
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/"
                    f"{urllib.parse.quote(clean, safe='/')}:/content"
                )
        else:
            _info, drive_id = self.resolve_drive()
            if "/" in file_url_or_guid:
                clean = _strip_library_prefix(file_url_or_guid)
                endpoint = (
                    f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/"
                    f"{urllib.parse.quote(clean, safe='/')}:/content"
                )
            else:
                endpoint = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{file_url_or_guid.strip('{}')}/content"

        data = request_json(
            endpoint,
            headers={"Authorization": f"Bearer {self.get_token()}", "Content-Type": "application/octet-stream"},
            method="PUT",
            data=local.read_bytes(),
            context=f"thay thế file bằng '{local.name}'",
        )
        item_id = data.get("id")
        versions = self.get_item_versions(drive_id, item_id) if item_id else []
        return {
            "status": "REPLACED",
            "name": data.get("name"),
            "size": data.get("size", local.stat().st_size),
            "id": item_id,
            "version": versions[0].get("id") if versions else "N/A",
            "webUrl": data.get("webUrl"),
            "modified": data.get("lastModifiedDateTime"),
        }

    def sync_folder_up(self, local_dir: str, target_folder: str, dry_run: bool = True) -> str:
        """Upload local files that are new or newer than their SharePoint copy.

        One-directional and non-destructive: nothing on SharePoint is ever
        deleted, and ``dry_run`` defaults to True so the plan is shown first.
        """
        local_root = Path(local_dir).expanduser().resolve()
        if not local_root.is_dir():
            raise Mcp365Error(f"Không tìm thấy thư mục local: {local_dir}", "Kiểm tra lại đường dẫn.")

        _info, drive_id = self.resolve_drive(target_folder if target_folder.startswith("http") else "")
        clean_folder = _strip_library_prefix(target_folder if not target_folder.startswith("http") else "")

        remote: dict[str, dict] = {}
        try:
            for item in self.list_folder_contents(drive_id, clean_folder):
                if "folder" not in item:
                    remote[item["name"]] = item
        except Mcp365Error as exc:
            raise Mcp365Error(
                f"Không đọc được thư mục đích trên SharePoint: {exc.message}", exc.remediation
            ) from exc

        planned: list[tuple[str, str, int]] = []
        for path in sorted(p for p in local_root.rglob("*") if p.is_file()):
            rel = path.relative_to(local_root).as_posix()
            match = remote.get(path.name)
            if match is None:
                planned.append((rel, "NEW", path.stat().st_size))
            elif path.stat().st_mtime > _iso_to_epoch(match.get("lastModifiedDateTime", "")):
                planned.append((rel, "NEWER", path.stat().st_size))

        if not planned:
            return f"✓ Đã đồng bộ: không có file nào trong `{local_dir}` cần tải lên."

        out = [
            f"# Đồng bộ `{local_dir}` → `{clean_folder or '/'}`",
            f"*{'CHẾ ĐỘ THỬ (dry run) — chưa tải gì lên' if dry_run else 'Đang tải lên'} · {len(planned)} file*\n",
            "| File | Trạng thái | Kích thước |",
            "| --- | --- | --- |",
        ]
        for rel, state, size in planned:
            out.append(f"| `{rel}` | {state} | {human_size(size)} |")

        if dry_run:
            out.append("\n> Gọi lại với `dry_run=False` để thực sự tải lên. Tool này KHÔNG bao giờ xoá file trên SharePoint.")
            return "\n".join(out)

        uploaded, failed = 0, []
        for rel, _state, _size in planned:
            try:
                sub = (Path(clean_folder) / Path(rel).parent).as_posix().strip("/.")
                self.upload_file(str(local_root / rel), sub or clean_folder, Path(rel).name)
                uploaded += 1
            except Mcp365Error as exc:
                failed.append(f"`{rel}`: {exc.message}")
        out.append(f"\n**Đã tải lên {uploaded}/{len(planned)} file.**")
        if failed:
            out.append("\n**Thất bại:**\n" + "\n".join(f"- {f}" for f in failed))
        return "\n".join(out)

    # --------------------------------------------------------------- search

    def search_files(self, query: str, max_results: int = 20, file_extension: str | None = None) -> list[dict[str, Any]]:
        cfg = get_config().sharepoint
        q = f"{query} path:{cfg.site_url}"
        if file_extension:
            q += f" fileextension:{file_extension.lstrip('.')}"

        url = (
            f"{cfg.site_url}/_api/search/query?querytext='{urllib.parse.quote(q)}'"
            f"&rowlimit={max_results}"
            f"&selectproperties='Title,Path,Author,Size,LastModifiedTime,UniqueId'"
        )
        data = request_json(url, headers=self._cookie_headers(), context=f"tìm kiếm '{query}' trên SharePoint")
        rows = (
            data.get("d", {})
            .get("query", {})
            .get("PrimaryQueryResult", {})
            .get("RelevantResults", {})
            .get("Table", {})
            .get("Rows", {})
            .get("results", [])
        )
        results = []
        for row in rows:
            cells = {c["Key"]: c["Value"] for c in row.get("Cells", {}).get("results", [])}
            path = cells.get("Path", "")
            if not path or path.endswith("/Forms/AllItems.aspx") or "/_catalogs/" in path:
                continue
            results.append(
                {
                    "title": cells.get("Title") or path.split("/")[-1],
                    "path": path,
                    "author": cells.get("Author", "Unknown"),
                    "size": int(cells.get("Size") or 0),
                    "modified": (cells.get("LastModifiedTime") or "")[:19].replace("T", " "),
                    "unique_id": (cells.get("UniqueId") or "").strip("{}"),
                }
            )
        return results

    # ----------------------------------------------------------- comparison

    @staticmethod
    def extract_lines(data: bytes, name: str) -> list[str] | None:
        """Extract comparable text lines from a document, or None if binary."""
        lower = name.lower()
        if lower.endswith(_WORD_EXTS):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as zf:
                    if "word/document.xml" in zf.namelist():
                        xml = zf.read("word/document.xml").decode("utf-8", errors="ignore")
                        lines = []
                        for para in re.findall(r"<w:p[ >].*?</w:p>", xml, re.DOTALL):
                            text = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", para)).strip()
                            if text:
                                lines.append(text)
                        return lines
            except (zipfile.BadZipFile, KeyError):
                return None
        elif lower.endswith(_TEXT_EXTS):
            return data.decode("utf-8", errors="ignore").splitlines()
        elif lower.endswith(_SHEET_EXTS):
            try:
                with zipfile.ZipFile(io.BytesIO(data)) as zf:
                    sheets = sorted(n.split("/")[-1] for n in zf.namelist() if n.startswith("xl/worksheets/"))
                    return [f"Sheet: {s}" for s in sheets]
            except (zipfile.BadZipFile, KeyError):
                return None
        return None

    @staticmethod
    def _render_diff(lines_a: list[str] | None, lines_b: list[str] | None, label_a: str, label_b: str, size_a: int, size_b: int, cap: int = 200) -> list[str]:
        if lines_a is None or lines_b is None:
            return [f"So sánh nhị phân: kích thước thay đổi {size_b - size_a:+d} bytes ({human_size(size_a)} → {human_size(size_b)})."]
        diff = list(difflib.unified_diff(lines_a, lines_b, fromfile=label_a, tofile=label_b, lineterm=""))
        if not diff:
            return [f"✓ Không có khác biệt nội dung giữa `{label_a}` và `{label_b}`."]
        adds = sum(1 for line in diff if line.startswith("+") and not line.startswith("+++"))
        dels = sum(1 for line in diff if line.startswith("-") and not line.startswith("---"))
        out = [f"### Tóm tắt thay đổi: **+{adds} thêm**, **-{dels} xoá**\n", "```diff", *diff[:cap]]
        if len(diff) > cap:
            out.append(f"... ({len(diff) - cap} dòng diff nữa được lược bớt)")
        out.append("```")
        return out

    def compare_versions(self, url_or_guid: str, version_a: str = "", version_b: str = "") -> str:
        cfg = get_config().sharepoint
        headers = self._cookie_headers()

        if url_or_guid.startswith("http"):
            file_rel = urllib.parse.unquote(urllib.parse.urlparse(url_or_guid).path)
            if "/sites/" not in file_rel:
                info = self.parse_sharepoint_url(url_or_guid)
                found = self.search_files(info.get("sourcedoc") or "", max_results=1) if info.get("sourcedoc") else []
                if not found:
                    raise Mcp365Error(f"Không xác định được tài liệu từ link: {url_or_guid}", "Thử truyền UniqueId (GUID).")
                file_rel = urllib.parse.unquote(urllib.parse.urlparse(found[0]["path"]).path)
        else:
            found = self.search_files(url_or_guid.strip("{}"), max_results=1)
            if not found:
                raise Mcp365Error(f"Không tìm thấy tài liệu cho GUID/đường dẫn: {url_or_guid}", "Dùng `search_sharepoint_files` để tìm.")
            file_rel = urllib.parse.unquote(urllib.parse.urlparse(found[0]["path"]).path)

        file_name = file_rel.split("/")[-1]
        versions_url = (
            f"{cfg.site_url}/_api/web/GetFileByServerRelativeUrl('{urllib.parse.quote(file_rel)}')/Versions"
        )
        data = request_json(versions_url, headers=headers, context=f"đọc lịch sử phiên bản của '{file_name}'")
        past = data.get("d", {}).get("results", [])
        if not past:
            return f"Không có phiên bản lịch sử nào cho `{file_name}`. Chỉ tồn tại bản hiện tại."

        target_a = None
        if version_a:
            target_a = next((v for v in past if v.get("VersionLabel") in (version_a, f"{version_a}.0")), None)
            if target_a is None:
                labels = ", ".join(v.get("VersionLabel", "?") for v in past)
                raise Mcp365Error(f"Không tìm thấy phiên bản '{version_a}'.", f"Các phiên bản có sẵn: {labels}")
        else:
            target_a = past[-1] if len(past) == 1 else past[-2]

        target_b = None
        if version_b and version_b.lower() not in ("latest", "current"):
            target_b = next((v for v in past if v.get("VersionLabel") in (version_b, f"{version_b}.0")), None)
            if target_b is None:
                labels = ", ".join(v.get("VersionLabel", "?") for v in past)
                raise Mcp365Error(f"Không tìm thấy phiên bản '{version_b}'.", f"Các phiên bản có sẵn: {labels}, latest")

        def fetch(version: dict | None) -> bytes:
            if version is None:
                url = f"https://{cfg.hostname}{urllib.parse.quote(file_rel, safe='/:')}"
            else:
                url = f"{cfg.site_url}/{urllib.parse.quote(version.get('Url', ''), safe='/:')}"
            return request_bytes(url, headers=headers, context="tải nội dung phiên bản")

        bytes_a, bytes_b = fetch(target_a), fetch(target_b)
        label_a = target_a.get("VersionLabel") if target_a else "Earlier"
        label_b = target_b.get("VersionLabel") if target_b else "Latest"

        out = [
            f"# 📊 So sánh phiên bản tài liệu: `{file_name}`",
            f"- **Phiên bản A:** `{label_a}` ({len(bytes_a):,} bytes)",
            f"- **Phiên bản B:** `{label_b}` ({len(bytes_b):,} bytes)\n",
            "---",
        ]
        out.extend(
            self._render_diff(
                self.extract_lines(bytes_a, file_name),
                self.extract_lines(bytes_b, file_name),
                f"Version {label_a}",
                f"Version {label_b}",
                len(bytes_a),
                len(bytes_b),
                cap=150,
            )
        )
        return "\n".join(out)

    def compare_documents(self, file_a: str, file_b: str) -> str:
        """Compare two documents, each either a local path or a SharePoint ref."""
        # A fresh temp dir per call: the previous implementation reused
        # /tmp/mcp_compare/file_a and picked glob("*")[0], so a leftover file
        # from an earlier comparison could silently win.
        tmp_root = Path(tempfile.mkdtemp(prefix="mcp365-compare-"))
        try:
            def resolve(target: str, tag: str) -> tuple[bytes, str]:
                path = Path(target).expanduser()
                if path.is_file():
                    return path.read_bytes(), path.name
                sub = tmp_root / tag
                sub.mkdir(parents=True, exist_ok=True)
                self.download_link(target, target_dir=str(sub))
                files = [p for p in sub.rglob("*") if p.is_file()]
                if not files:
                    raise Mcp365Error(f"Không tải được tài liệu: {target}", "Kiểm tra link hoặc quyền truy cập.")
                return files[0].read_bytes(), files[0].name

            data_a, name_a = resolve(file_a, "file_a")
            data_b, name_b = resolve(file_b, "file_b")

            out = [
                "# 📊 Báo cáo so sánh tài liệu",
                f"- **File A:** `{name_a}` ({len(data_a):,} bytes)",
                f"- **File B:** `{name_b}` ({len(data_b):,} bytes)\n",
                "---",
            ]
            out.extend(
                self._render_diff(
                    self.extract_lines(data_a, name_a),
                    self.extract_lines(data_b, name_b),
                    f"File A ({name_a})",
                    f"File B ({name_b})",
                    len(data_a),
                    len(data_b),
                )
            )
            return "\n".join(out)
        finally:
            shutil.rmtree(tmp_root, ignore_errors=True)


def _iso_to_epoch(value: str) -> float:
    from datetime import datetime

    try:
        return datetime.fromisoformat((value or "").replace("Z", "+00:00")).timestamp()
    except ValueError:
        return 0.0
