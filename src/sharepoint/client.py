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
import logging
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
    ConcurrentEditError,
    Mcp365Error,
    UnsupportedOperationError,
)
from common.http import capture_cookie, request, request_bytes, request_json

logger = logging.getLogger(__name__)


def _extract_host_from_path(path: str, default_host: str) -> str:
    if path.startswith("/sites/"):
        parts = path[len("/sites/"):].split("/")
        if parts:
            segment = parts[0].split(":")[0]
            if "," in segment:
                segment = segment.split(",")[0]
            if "." in segment and not any(c in segment for c in " ,?#"):
                return segment.lower()
    return default_host.lower()

def _both_channels_failed(what: str, cookie_error: Exception, graph_error: Exception) -> Mcp365Error:
    """Report both channels, so the real cause (often the cookie one) is not hidden."""

    def text(exc: Exception) -> str:
        return getattr(exc, "message", "") or str(exc) or type(exc).__name__

    remedies = [r for r in (getattr(cookie_error, "remediation", ""), getattr(graph_error, "remediation", "")) if r]
    return Mcp365Error(
        f"Cả 2 kênh SharePoint đều lỗi khi {what}.\n"
        f"• Kênh chính (cookie Chrome): {text(cookie_error)}\n"
        f"• Kênh phụ (Graph qua Azure CLI): {text(graph_error)}",
        " | ".join(dict.fromkeys(remedies)),
    )


def _odata_literal(value: str) -> str:
    """``value`` as the inside of an OData string literal in a URL path (``'`` doubled)."""
    return urllib.parse.quote(value.replace("'", "''"), safe="/")


_WRITE_METHODS = ("POST", "PUT", "DELETE", "PATCH")
_DRIVE_PATH_RE = re.compile(r"^/drives/([^/?:]+)")
_GUID_RE = re.compile(r"([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})")
#: Every kind of SharePoint site collection: team/communication sites under
#: ``/sites/`` or ``/teams/``, and personal OneDrives under ``/personal/``.
#: Only ``/sites/`` used to be recognised, so a OneDrive link (every file shared
#: in a Teams chat lives there) was silently resolved against the configured
#: site and failed with "Requested site could not be found".
_SITE_RE = re.compile(r"/(sites|teams|personal)/([^/?#]+)", re.IGNORECASE)
#: Default document library of a site: ``Shared Documents`` on team sites,
#: ``Documents`` on a personal OneDrive.
_SHARED_DOCS_RE = re.compile(r"^/(?:sites|teams|personal)/[^/]+/(?:Shared Documents|Documents)/?", re.IGNORECASE)
#: Sharing links look like ``/:x:/r/...`` or ``/:u:/g/...`` - the letter is the
#: file kind (x=Excel, w=Word, p=PowerPoint, b=PDF, u=any other file, ...).
_SHARING_LINK_RE = re.compile(r"/:([a-z]):/", re.IGNORECASE)

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
        #: host -> (FedAuth, issued_at) minted from rtFa; see _mint_fed_auth().
        self._minted: dict[str, tuple[str, float]] = {}
        #: drive id -> URL of the site that owns it (``https://host/sites/X``).
        self._drive_sites: dict[str, str] = {}
        #: site URL -> (FormDigestValue, expires_at) for state-changing requests.
        self._form_digests: dict[str, tuple[str, float]] = {}
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

    def _get_form_digest(self, site_url: str = "") -> str:
        """FormDigestValue for writes to ``site_url``, cached per site.

        A digest is only accepted by the site that issued it. It used to be
        requested from ``https://{host}/_api/contextinfo`` - the host's *root*
        site - so every cookie write into ``/sites/X`` was refused (403 creating
        a folder, 401 uploading) however valid the session was.
        """
        site_url = (site_url or get_config().sharepoint.site_url).rstrip("/")
        key = site_url.lower()
        now = time.time()
        cached = self._form_digests.get(key)
        if cached and now < (cached[1] - 60):
            return cached[0]
        headers = self._cookie_headers(
            accept="application/json;odata=verbose", host=urllib.parse.urlparse(site_url).netloc
        )
        res = request_json(
            f"{site_url}/_api/contextinfo",
            method="POST",
            headers=headers,
            context=f"xin FormDigest cho {site_url}",
        )
        info = res.get("d", {}).get("GetContextWebInformation", {})
        digest = info.get("FormDigestValue", "")
        if not digest:
            raise Mcp365Error(
                f"SharePoint không trả FormDigest cho {site_url}, nên không ghi được qua cookie.",
                f"Mở {site_url} trong Chrome (đăng nhập, tick 'Stay signed in') rồi thử lại.",
            )
        timeout = float(info.get("FormDigestTimeoutSeconds", 1800))
        self._form_digests[key] = (digest, now + timeout)
        return digest

    def _site_url_for(self, path: str) -> str:
        """URL of the site owning the drive in ``/drives/{id}/...``, if resolve_drive saw it."""
        match = _DRIVE_PATH_RE.match(path)
        return self._drive_sites.get(match.group(1), "") if match else ""

    def _cookie_then_graph(self, what: str, cookie_call, graph_call):
        """Run the cookie channel, fall back to Graph, and report both if both fail.

        No fallback on ``ConcurrentEditError`` (409/412/423): the server reached
        the file and refused the write, and Graph serves the same file. Falling
        back cost an ``az`` call and, whenever Graph was itself broken (CAE),
        replaced "file changed, nothing written" with an unrelated 401.

        403 and 404 do fall back: the cookie channel is scoped to one host and
        site - a path with no host goes to the configured one, which cannot serve
        a drive on another host (every OneDrive) - and its 403s are often
        cookie-specific (non-persistent FedAuth, error 917656). Graph shares
        neither. The cookie error is not lost: ``_both_channels_failed`` shows it.
        """
        cookie_error: Mcp365Error | None = None
        if cookie_call is not None:
            try:
                return cookie_call()
            except ConcurrentEditError:
                raise
            except Mcp365Error as exc:
                cookie_error = exc
                logger.info("SharePoint cookie channel failed (%s); falling back to Azure CLI Graph", exc)
        try:
            return graph_call()
        except ConcurrentEditError:
            raise
        except Mcp365Error as graph_exc:
            if cookie_error is not None:
                raise _both_channels_failed(what, cookie_error, graph_exc) from graph_exc
            raise

    def _graph_json(
        self,
        clean_path: str,
        method: str = "GET",
        payload: bytes | None = None,
        extra_headers: dict[str, str] | None = None,
        json_body: bool = False,
        context: str = "",
    ) -> dict[str, Any]:
        """One Graph request with the Azure CLI token, re-minted once on 401."""
        url = f"https://graph.microsoft.com/v1.0{clean_path}"

        def send() -> dict[str, Any]:
            hdrs = {"Authorization": f"Bearer {self.get_token()}", "Accept": "application/json"}
            if extra_headers:
                hdrs.update(extra_headers)
            if json_body and "Content-Type" not in hdrs:
                hdrs["Content-Type"] = "application/json"
            return request_json(
                url, headers=hdrs, method=method, data=payload, context=context or f"gọi Graph (fallback) {clean_path}"
            )

        try:
            return send()
        except AuthExpiredError:
            self._token = None
            self._token_expires = 0.0
            self.get_token(force_refresh=True)
            return send()

    def call_sharepoint_or_graph(
        self,
        path: str,
        method: str = "GET",
        body: dict | None = None,
        data: bytes | None = None,
        host: str = "",
        extra_headers: dict[str, str] | None = None,
        context: str = "",
    ) -> dict[str, Any]:
        """Execute a drive or site operation: cookie-first on SharePoint, Azure CLI Graph fallback.

        Writes on a drive seen by :meth:`resolve_drive` go to that drive's own
        site (``{site}/_api/v2.0``) with that site's FormDigest. Fallback rules:
        :meth:`_cookie_then_graph`.
        """
        clean_path = path
        if clean_path.startswith("https://graph.microsoft.com/v1.0"):
            clean_path = clean_path[len("https://graph.microsoft.com/v1.0"):]
        if not clean_path.startswith("/"):
            clean_path = f"/{clean_path}"

        cfg_host = get_config().sharepoint.hostname.lower()
        target_host = (host or _extract_host_from_path(clean_path, cfg_host)).lower()
        payload = json.dumps(body).encode("utf-8") if body is not None else data
        has_composite_site_id = (
            clean_path.startswith("/sites/")
            and len(clean_path.split("/")) > 2
            and "," in clean_path.split("/")[2]
        )

        # 1. Primary channel: SharePoint native _api/v2.0 using browser cookies
        def cookie_call() -> dict[str, Any]:
            is_write = method in _WRITE_METHODS
            site_url = (self._site_url_for(clean_path) if is_write else "") or f"https://{target_host}"
            headers = self._cookie_headers(accept="application/json", host=urllib.parse.urlparse(site_url).netloc)
            if extra_headers:
                headers.update(extra_headers)
            if is_write and "X-RequestDigest" not in headers:
                # Without a digest the write is refused anyway, with a bare 403
                # that hides why - so a digest failure is the error to report.
                headers["X-RequestDigest"] = self._get_form_digest(site_url)
            if body is not None and "Content-Type" not in headers:
                headers["Content-Type"] = "application/json"
            return request_json(
                f"{site_url}/_api/v2.0{clean_path}",
                headers=headers,
                method=method,
                data=payload,
                context=context or f"gọi SharePoint REST {clean_path}",
            )

        # 2. Fallback channel: Microsoft Graph via Azure CLI token
        def graph_call() -> dict[str, Any]:
            return self._graph_json(clean_path, method, payload, extra_headers, body is not None, context)

        return self._cookie_then_graph(
            context or clean_path, None if has_composite_site_id else cookie_call, graph_call
        )

    def call_graph(self, path: str, method: str = "GET", body: dict | None = None, context: str = "") -> dict:
        return self.call_sharepoint_or_graph(path=path, method=method, body=body, context=context)
    # --------------------------------------------------------- cookie auth

    def _cookie_headers(self, accept: str = "application/json;odata=verbose", host: str = "") -> dict[str, str]:
        """Headers for the direct-session channel, scoped to one host.

        Cookie order (``rtFa`` first) and a browser User-Agent are both required;
        the previous code applied them only on the download path, so search and
        version history went out malformed.

        ``FedAuth`` is issued **per host**: ``tenant.sharepoint.com`` and
        ``tenant-my.sharepoint.com`` each carry their own. The old lookup matched
        ``%sharepoint.com%`` and took the most recently used row - whichever host
        the browser touched last - so OneDrive requests could leave with the
        team-site cookie and vice versa. ``rtFa`` is tenant-wide.
        """
        host = (host or get_config().sharepoint.hostname).lower()
        fed_auth = ChromeCookieDecryptor.get_cookies_for_domain(host, ["FedAuth"]).get("FedAuth")
        rt_fa = ChromeCookieDecryptor.get_cookies_for_domain(host, ["rtFa"]).get("rtFa") or (
            ChromeCookieDecryptor.get_cookies_for_domain("sharepoint.com", ["rtFa"]).get("rtFa")
        )
        if not fed_auth and rt_fa:
            fed_auth = self._mint_fed_auth(host, rt_fa)
        if not fed_auth or not rt_fa:
            missing = " và ".join(n for n, v in (("rtFa", rt_fa), ("FedAuth", fed_auth)) if not v)
            raise AuthExpiredError(
                f"Thiếu cookie {missing} cho host '{host}', và không xin được FedAuth từ rtFa.",
                "rtFa là cookie đăng nhập chung của tenant: thiếu nó nghĩa là phiên SharePoint trong Chrome "
                "đã hết - đăng nhập lại và tick 'Stay signed in'. Nếu chỉ thiếu FedAuth, mở "
                f"https://{host} trong Chrome rồi thử lại.",
            )
        return {
            "Cookie": f"rtFa={rt_fa}; FedAuth={fed_auth}",
            "User-Agent": get_config().http.user_agent,
            "Accept": accept,
        }

    def _mint_fed_auth(self, host: str, rt_fa: str) -> str:
        """Obtain a host-scoped ``FedAuth`` from the tenant-wide ``rtFa``.

        Why this exists: a user can be happily browsing OneDrive while Chrome's
        cookie database holds no ``FedAuth`` for ``tenant-my.sharepoint.com`` at
        all - when it is a *session* cookie, Chrome keeps it in memory only.
        ``rtFa`` is persistent, and SharePoint trades it for a per-host
        ``FedAuth`` through a redirect hand-off; this replays that hand-off,
        which is exactly what the browser does on first visit to a new host.
        Cached for 30 minutes.
        """
        cached = self._minted.get(host)
        if cached and time.time() - cached[1] < 1800:
            return cached[0]
        headers = {"Cookie": f"rtFa={rt_fa}", "Accept": "text/html,*/*"}
        for path in ("/_forms/default.aspx?wa=wsignin1.0", "/"):
            value = capture_cookie(f"https://{host}{path}", headers=headers, name="FedAuth", host=host)
            if value:
                self._minted[host] = (value, time.time())
                return value
        return ""

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

        site_match = _SITE_RE.search(path)
        if site_match:
            kind, name = site_match.group(1).lower(), site_match.group(2)
            info["site_name"] = name
            info["site_path"] = f"/{kind}/{name}"
            info["is_personal"] = info["is_personal"] or kind == "personal"
        elif info["hostname"].lower() == cfg.hostname.lower():
            # Same tenant host but no site segment: the configured site is the
            # best guess (legacy links such as bare Doc.aspx?sourcedoc=...).
            info["site_name"] = cfg.site_name
            info["site_path"] = cfg.site_path
        else:
            # A different host with no site segment is that host's root site.
            # Never borrow the configured site path for a foreign host.
            info["site_name"] = ""
            info["site_path"] = ""

        if "id" in query:
            info["folder_path"] = query["id"][0]
            if not info["sourcedoc"]:
                info["type"] = "folder"
        elif not info["folder_path"] and not info["sourcedoc"]:
            clean_rel = _strip_library_prefix(path)
            if clean_rel:
                if path.lower().endswith(_BINARY_EXTS):
                    info["type"] = "document"
                    info["file_name"] = Path(clean_rel).name
                else:
                    info["type"] = "folder"
                    info["folder_path"] = clean_rel
        if info["type"] == "unknown":
            kind_letter = _SHARING_LINK_RE.search(path)
            if kind_letter:
                info["type"] = "folder" if kind_letter.group(1).lower() == "f" else "document"
            elif path.lower().endswith(_BINARY_EXTS):
                info["type"] = "document"
        return info

    def get_site_id(self, hostname: str, site_path: str) -> str:
        key = f"{hostname}:{site_path}"
        if key not in self._site_cache:
            # ``/sites/{host}`` is the root site; ``/sites/{host}:{path}`` any other.
            endpoint = f"/sites/{hostname}:{site_path}" if site_path else f"/sites/{hostname}"
            data = self.call_graph(endpoint, context=f"tra cứu site {site_path or hostname}")
            self._site_cache[key] = data["id"]
        return self._site_cache[key]

    def get_default_drive_id(self, site_id: str) -> str:
        if site_id not in self._drive_cache:
            # ``/drive`` (singular) is the site's default library. ``/drives``[0]
            # was merely the first one listed, which on a site with several
            # libraries is not necessarily "Documents".
            drive = self.call_graph(f"/sites/{site_id}/drive", context="tra cứu thư viện tài liệu mặc định")
            if not drive.get("id"):
                raise Mcp365Error(f"Site {site_id} không có thư viện tài liệu nào.", "Kiểm tra lại site đích.")
            self._drive_cache[site_id] = drive["id"]
        return self._drive_cache[site_id]

    def _drive_web_url(self, drive_id: str) -> str:
        """Root URL of a document library, whatever it is called.

        Used to build a direct file URL from Graph metadata instead of assuming
        the library is named ``Shared Documents`` - which is false for every
        personal OneDrive (``Documents``) and for any custom library.
        """
        key = f"web:{drive_id}"
        if key not in self._drive_cache:
            drive = self.call_graph(f"/drives/{drive_id}?$select=webUrl", context="đọc URL thư viện tài liệu")
            self._drive_cache[key] = drive.get("webUrl", "").rstrip("/")
        return self._drive_cache[key]

    def _item_file_url(self, drive_id: str, item: dict[str, Any]) -> str:
        """Direct URL of a drive item's bytes.

        Prefers Graph's pre-authenticated ``@microsoft.graph.downloadUrl``; falls
        back to the library root plus the item's parent path.
        """
        pre_authed = item.get("@microsoft.graph.downloadUrl")
        if pre_authed:
            return pre_authed
        base = self._drive_web_url(drive_id)
        parent = urllib.parse.unquote(item.get("parentReference", {}).get("path", "").split("root:")[-1])
        return f"{base}{urllib.parse.quote(parent, safe='/')}/{urllib.parse.quote(item['name'])}"

    # ------------------------------------------------- single-file round trip

    def resolve_file(self, url_or_guid: str) -> tuple[str, dict[str, Any]]:
        """Graph metadata for one file as ``(drive_id, item)``.

        The item carries ``eTag`` - needed to write back without clobbering
        someone else's edit - and usually a pre-authenticated download URL.
        """
        if url_or_guid.startswith("http"):
            info, drive_id = self.resolve_drive(url_or_guid)
            if info.get("sourcedoc"):
                return drive_id, self.get_item_by_guid(drive_id, info["sourcedoc"])
            rel = _strip_library_prefix(urllib.parse.unquote(urllib.parse.urlparse(url_or_guid).path))
        else:
            _info, drive_id = self.resolve_drive()
            guid = _GUID_RE.search(url_or_guid)
            if guid:
                return drive_id, self.get_item_by_guid(drive_id, guid.group(1))
            rel = _strip_library_prefix(url_or_guid)
        encoded_rel = urllib.parse.quote(rel, safe="/")
        endpoint = f"/drives/{drive_id}/root:/{encoded_rel}:" if encoded_rel else f"/drives/{drive_id}/root"
        item = self.call_graph(endpoint, context=f"tra cứu file '{rel}'")
        return drive_id, item

    def read_file_bytes(self, drive_id: str, item: dict[str, Any]) -> bytes:
        url = self._item_file_url(drive_id, item)
        return request_bytes(url, headers=self._download_headers(url), context=f"tải '{item.get('name', '')}'")

    def put_file_bytes(self, drive_id: str, item_id: str, data: bytes, if_match: str = "") -> dict[str, Any]:
        """Upload new content as a new version, optionally guarded by ``If-Match``.

        With ``if_match`` set to the eTag seen at read time, Graph/SharePoint answers 412
        if anyone saved in between; that surfaces as ``ConcurrentEditError``
        and nothing is written.
        """
        headers = {"Content-Type": "application/octet-stream"}
        if if_match:
            headers["If-Match"] = if_match
        return self.call_sharepoint_or_graph(
            f"/drives/{drive_id}/items/{item_id}/content",
            method="PUT",
            data=data,
            extra_headers=headers,
            context="ghi phiên bản mới lên SharePoint",
        )
    def resolve_drive(self, url: str = "") -> tuple[dict[str, Any], str]:
        """Return ``(url_info, drive_id)`` for a URL, falling back to config."""
        info = self.parse_sharepoint_url(url) if url.startswith("http") else self._default_info()
        hostname = info["hostname"]
        site_path = info["site_path"] or ""
        key = f"{hostname}:{site_path}"
        if key in self._drive_cache:
            return info, self._drive_cache[key]
        what = f"tra cứu thư viện {site_path or hostname}"

        # 1. Primary: Direct lookup via SharePoint _api/v2.0/drive with cookies
        def by_cookie() -> str:
            headers = self._cookie_headers(accept="application/json", host=hostname)
            data = request_json(f"https://{hostname}{site_path}/_api/v2.0/drive", headers=headers, context=what)
            if not data.get("id"):
                raise Mcp365Error(f"SharePoint không trả id thư viện khi {what}.")
            if data.get("webUrl"):
                self._drive_cache[f"web:{data['id']}"] = data["webUrl"].rstrip("/")
            return data["id"]

        # 2. Fallback: Graph site ID resolution + /sites/{site_id}/drive. Its
        #    error used to be the only one reported, hiding the cookie cause.
        def by_graph() -> str:
            return self.get_default_drive_id(self.get_site_id(info["hostname"], info["site_path"]))

        drive_id = self._cookie_then_graph(what, by_cookie, by_graph)
        self._drive_cache[key] = drive_id
        self._drive_sites[drive_id] = f"https://{hostname}{site_path}"
        return info, drive_id

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

    def _download_headers(self, file_url: str) -> dict[str, str]:
        """Cookie headers for the host that actually serves ``file_url``.

        Graph's pre-authenticated download URLs carry their own ``tempauth``
        token, so they get no cookie at all.
        """
        if "tempauth=" in file_url:
            return {"User-Agent": get_config().http.user_agent, "Accept": "*/*"}
        return self._cookie_headers(accept="*/*", host=urllib.parse.urlparse(file_url).netloc)

    def _download_to(self, file_url: str, dest: Path, results: list[tuple[str, int, str]]) -> bool:
        dest.parent.mkdir(parents=True, exist_ok=True)
        try:
            data = request_bytes(file_url, headers=self._download_headers(file_url), context=f"tải '{dest.name}'")
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
        results: list[tuple[str, int, str]] = []

        if url_or_guid.startswith("http"):
            self._download_from_url(url_or_guid, target_path, results)
        elif _GUID_RE.search(url_or_guid):
            # A bare GUID names no site, so it can only be looked up in the
            # configured one. Pass a full URL for files anywhere else.
            guid = _GUID_RE.search(url_or_guid).group(1)
            _info, drive_id = self.resolve_drive()
            item = self.get_item_by_guid(drive_id, guid)
            self._download_to(self._item_file_url(drive_id, item), target_path / item["name"], results)
        else:
            raise Mcp365Error(
                f"Không nhận dạng được link hoặc GUID SharePoint: {url_or_guid}",
                "Truyền URL đầy đủ (https://...) hoặc UniqueId dạng GUID.",
            )
        return self._render_download_report(results, str(target_path))

    def _download_from_url(self, url: str, target_path: Path, results: list[tuple[str, int, str]]) -> None:
        parsed = urllib.parse.urlparse(url)
        parsed_path = urllib.parse.unquote(parsed.path)
        sharing = _SHARING_LINK_RE.search(parsed.path)

        # 1. Personal OneDrive sharing link to an Office file -> WOPI context.
        if "-my.sharepoint.com" in url and sharing and sharing.group(1).lower() in "xwpb":
            headers = self._cookie_headers(accept="*/*", host=parsed.netloc)
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
            self._download_to(file_url, target_path / (wopi.get("FileName") or "downloaded_file"), results)
            return

        # 2. Direct path to a file in any site kind (/sites, /teams, /personal).
        #    NOTE: this branch previously fell through into the Graph folder
        #    logic because `drive_id = ...` was indented one level too far out,
        #    raising UnboundLocalError *after* the file had been written. It also
        #    only accepted /sites/, so OneDrive paths - every Teams chat
        #    attachment - fell through to Graph and 404'd.
        if parsed_path.lower().endswith(_BINARY_EXTS) and _SITE_RE.search(parsed_path):
            file_url = f"https://{parsed.netloc}{urllib.parse.quote(parsed_path, safe='/:')}"
            self._download_to(file_url, target_path / parsed_path.split("/")[-1], results)
            return

        # 3. Any other file sharing link (``:u:`` zip/json/..., ``:i:``, ``:v:``,
        #    or Office links on a team site): SharePoint serves the bytes when
        #    asked with ``download=1``. Folder links (``:f:``) go to Graph below.
        if sharing and sharing.group(1).lower() != "f":
            sep = "&" if parsed.query else "?"
            name = urllib.parse.parse_qs(parsed.query).get("file", [""])[0] or "downloaded_file"
            self._download_to(f"{url}{sep}download=1", target_path / name, results)
            return

        # 4. Anything else -> resolve through Graph (single document or folder).
        info, drive_id = self.resolve_drive(url)
        if info["type"] == "document" and info["sourcedoc"]:
            item = self.get_item_by_guid(drive_id, info["sourcedoc"])
            self._download_to(self._item_file_url(drive_id, item), target_path / item["name"], results)
            return

        clean = _strip_library_prefix(info.get("folder_path") or "")
        endpoint = f"/drives/{drive_id}/root:/{urllib.parse.quote(clean)}" if clean else f"/drives/{drive_id}/root"
        folder_item = self.call_graph(endpoint, context="mở thư mục SharePoint")
        self._sync_folder_down(drive_id, folder_item["id"], target_path, results)

    def _sync_folder_down(
        self,
        drive_id: str,
        folder_id: str,
        local_dir: Path,
        results: list[tuple[str, int, str]],
        skip_large_media: bool = True,
    ) -> None:
        items = self.call_graph(f"/drives/{drive_id}/items/{folder_id}/children", context="liệt kê nội dung thư mục").get(
            "value", []
        )
        for item in items:
            name = item["name"]
            if "folder" in item:
                self._sync_folder_down(drive_id, item["id"], local_dir / name, results)
                continue
            if skip_large_media and name.lower().endswith((".mp4", ".mov")) and item.get("size", 0) > 50 * 1024 * 1024:
                results.append((name, 0, "Bỏ qua: video lớn hơn 50MB"))
                continue
            self._download_to(self._item_file_url(drive_id, item), local_dir / name, results)

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
        results: list[tuple[str, int, str]] = []
        for video in videos:
            # Recordings usually sit in the organiser's OneDrive (-my host), so
            # the cookie must be picked per file rather than once for the site.
            parsed = urllib.parse.urlparse(video["path"])
            file_url = f"https://{parsed.netloc}{urllib.parse.quote(urllib.parse.unquote(parsed.path), safe='/:')}"
            self._download_to(file_url, target_path / video["title"], results)
        return self._render_download_report(results, str(target_path))

    # -------------------------------------------------------------- uploads

    def _folder_exists(self, drive_id: str, path: str) -> bool:
        """True if ``path`` is a folder; False only when a channel answered 404.

        Any other failure (expired session, 403) is raised. Treating it as
        "missing" led to a POST that failed with a second, misleading error.
        """
        try:
            item = self.call_graph(
                f"/drives/{drive_id}/root:/{urllib.parse.quote(path, safe='/')}",
                context=f"kiểm tra thư mục '{path}'",
            )
        except Mcp365Error as exc:
            if "HTTP 404" in str(exc):
                return False
            raise
        return "folder" in item

    def ensure_folder(self, drive_id: str, folder_path: str) -> None:
        parts = [p for p in _strip_library_prefix(folder_path).split("/") if p]
        # Check before creating: users often may upload into a folder but not
        # create siblings at the library root, so a blind POST answers 403.
        if not parts or self._folder_exists(drive_id, "/".join(parts)):
            return
        current, missing = "", False
        for index, part in enumerate(parts):
            path = f"{current}/{part}" if current else part
            # The full path is known to be missing, and so is everything under
            # a missing folder: only the leading, existing folders cost a GET.
            missing = missing or index == len(parts) - 1 or not self._folder_exists(drive_id, path)
            if missing:
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
            current = path

    def _add_file_via_cookies(self, drive_id: str, folder: str, file_name: str, content: bytes) -> dict[str, Any]:
        """Upload through REST v1 ``Files/add`` on the library's own site.

        This is the cookie upload route verified live on the tenant (``PUT
        _api/v2.0/.../content`` answered 401 there, with the host-root digest;
        v2.0 writes with the site digest were not re-tested). Returns the
        Graph-shaped fields :meth:`upload_file` reads.
        """
        site_url = self._drive_sites.get(drive_id, "")
        if not site_url:
            raise Mcp365Error(f"Chưa biết site chứa drive {drive_id} để upload qua cookie.")
        library = urllib.parse.unquote(urllib.parse.urlparse(self._drive_web_url(drive_id)).path).rstrip("/")
        if not library:
            raise Mcp365Error(f"Không đọc được đường dẫn thư viện của drive {drive_id}.")
        host = urllib.parse.urlparse(site_url).netloc
        headers = self._cookie_headers(accept="application/json;odata=verbose", host=host)
        headers["X-RequestDigest"] = self._get_form_digest(site_url)
        headers["Content-Type"] = "application/octet-stream"
        folder_url = f"{library}/{folder}" if folder else library
        res = request_json(
            f"{site_url}/_api/web/GetFolderByServerRelativeUrl('{_odata_literal(folder_url)}')"
            f"/Files/add(url='{_odata_literal(file_name)}',overwrite=true)",
            headers=headers,
            method="POST",
            data=content,
            context=f"upload '{file_name}'",
        )
        added = res.get("d", {})
        server_url = added.get("ServerRelativeUrl", "")
        return {
            "name": added.get("Name", file_name),
            "size": int(added.get("Length") or len(content)),
            "id": added.get("UniqueId"),
            "webUrl": f"https://{host}{urllib.parse.quote(server_url, safe='/')}" if server_url else None,
        }

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
            content = local.read_bytes()
            data = self._cookie_then_graph(
                f"upload '{file_name}'",
                lambda: self._add_file_via_cookies(drive_id, clean_folder, file_name, content),
                lambda: self._graph_json(
                    f"/drives/{drive_id}/root:/{encoded}:/content",
                    method="PUT",
                    payload=content,
                    extra_headers={"Content-Type": "application/octet-stream"},
                    context=f"upload '{file_name}'",
                ),
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
                endpoint = f"/drives/{drive_id}/items/{info['sourcedoc']}/content"
            else:
                clean = _strip_library_prefix(info.get("folder_path") or info.get("file_name") or "")
                endpoint = f"/drives/{drive_id}/root:/{urllib.parse.quote(clean, safe='/')}:/content"
        else:
            _info, drive_id = self.resolve_drive()
            if "/" in file_url_or_guid:
                clean = _strip_library_prefix(file_url_or_guid)
                endpoint = f"/drives/{drive_id}/root:/{urllib.parse.quote(clean, safe='/')}:/content"
            else:
                endpoint = f"/drives/{drive_id}/items/{file_url_or_guid.strip('{}')}/content"

        data = self.call_sharepoint_or_graph(
            endpoint,
            method="PUT",
            data=local.read_bytes(),
            extra_headers={"Content-Type": "application/octet-stream"},
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

    # -------------------------------------------------------------- delete

    def _server_relative_url(self, drive_id: str, item: dict[str, Any]) -> str:
        """Server-relative URL of a drive item, e.g. ``/sites/X/Shared Documents/a/b``.

        Built from the library root and ``parentReference.path``, not from
        ``webUrl``: an Office file's webUrl is a ``Doc.aspx?sourcedoc=`` link.
        """
        library = urllib.parse.unquote(urllib.parse.urlparse(self._drive_web_url(drive_id)).path).rstrip("/")
        if not library:
            raise Mcp365Error(f"Không đọc được đường dẫn thư viện của drive {drive_id}.")
        parent = urllib.parse.unquote(item.get("parentReference", {}).get("path", "").split("root:")[-1]).rstrip("/")
        return f"{library}{parent}/{item['name']}"

    def describe_item(self, url_or_guid: str) -> dict[str, Any]:
        """What a delete would remove: name, path, size and whether it is a folder."""
        drive_id, item = self.resolve_file(url_or_guid)
        if not item.get("id") or not item.get("name"):
            raise Mcp365Error(f"Không tìm thấy item SharePoint: {url_or_guid}", "Kiểm tra lại URL.")
        folder = item.get("folder")
        return {
            "drive_id": drive_id,
            "id": item["id"],
            "name": item["name"],
            "path": self._server_relative_url(drive_id, item),
            "size": int(item.get("size") or 0),
            "is_folder": folder is not None,
            "child_count": (folder or {}).get("childCount"),
        }

    def delete_item(self, target: dict[str, Any], permanent: bool = False) -> dict[str, Any]:
        """Delete the file or folder (with everything in it) that :meth:`describe_item` returned.

        ``permanent=False`` moves it to the site Recycle Bin: restorable, but
        it still counts against the site quota. ``permanent=True`` deletes it
        for good, the only way to give space back on a full site (HTTP 507)
        without a site admin emptying the second-stage bin.

        REST v1 on the library's own site with that site's digest - the same
        cookie route :meth:`_add_file_via_cookies` verified on the tenant.
        ``recycle()`` is the Recycle Bin move; ``X-HTTP-Method: DELETE`` is
        ``DeleteObject``, which bypasses the bin.
        """
        drive_id = target["drive_id"]
        site_url = self._drive_sites.get(drive_id, "")
        if not site_url:
            raise Mcp365Error(f"Chưa biết site chứa drive {drive_id} để xoá qua cookie.")
        kind = "Folder" if target["is_folder"] else "File"
        endpoint = f"{site_url}/_api/web/Get{kind}ByServerRelativeUrl('{_odata_literal(target['path'])}')"
        headers = self._cookie_headers(
            accept="application/json;odata=verbose", host=urllib.parse.urlparse(site_url).netloc
        )
        headers["X-RequestDigest"] = self._get_form_digest(site_url)
        if permanent:
            headers.update({"X-HTTP-Method": "DELETE", "IF-MATCH": "*"})
            request_json(endpoint, method="POST", headers=headers, context=f"xoá vĩnh viễn '{target['name']}'")
        else:
            request_json(
                f"{endpoint}/recycle()", method="POST", headers=headers, context=f"chuyển '{target['name']}' vào thùng rác"
            )
        return {**target, "permanent": permanent}

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
        host = cfg.hostname

        def locate(file_url: str) -> tuple[str, str]:
            parsed = urllib.parse.urlparse(file_url)
            return parsed.netloc or cfg.hostname, urllib.parse.unquote(parsed.path)

        if url_or_guid.startswith("http"):
            host, file_rel = locate(url_or_guid)
            if not (_SITE_RE.search(file_rel) and file_rel.lower().endswith(_BINARY_EXTS)):
                info = self.parse_sharepoint_url(url_or_guid)
                found = self.search_files(info.get("sourcedoc") or "", max_results=1) if info.get("sourcedoc") else []
                if not found:
                    raise Mcp365Error(f"Không xác định được tài liệu từ link: {url_or_guid}", "Thử truyền UniqueId (GUID).")
                host, file_rel = locate(found[0]["path"])
        else:
            found = self.search_files(url_or_guid.strip("{}"), max_results=1)
            if not found:
                raise Mcp365Error(f"Không tìm thấy tài liệu cho GUID/đường dẫn: {url_or_guid}", "Dùng `search_sharepoint_files` để tìm.")
            host, file_rel = locate(found[0]["path"])

        # The versions API and the bytes both live on the file's *own* site.
        # These used to go to the configured site whatever the file's location,
        # so comparing anything outside it (a OneDrive file, another team site)
        # asked VF_AIDV about a path it does not own.
        site = _SITE_RE.search(file_rel)
        site_url = f"https://{host}{site.group(0) if site else ''}"
        headers = self._cookie_headers(host=host)

        file_name = file_rel.split("/")[-1]
        versions_url = f"{site_url}/_api/web/GetFileByServerRelativeUrl('{urllib.parse.quote(file_rel)}')/Versions"
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
                url = f"https://{host}{urllib.parse.quote(file_rel, safe='/:')}"
            else:
                url = f"{site_url}/{urllib.parse.quote(version.get('Url', ''), safe='/:')}"
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
