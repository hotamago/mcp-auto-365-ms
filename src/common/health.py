"""End-to-end diagnostics for every authentication channel.

Motivation: with the old code a broken session surfaced as the single line
``Error searching SharePoint files: HTTP Error 403: Forbidden``, which does not
distinguish "cookie expired" from "wrong site" from "missing permission". Every
probe below reports the channel, the evidence and the concrete fix.
"""

from __future__ import annotations

import subprocess
import time
from datetime import datetime
from typing import Any

from .chrome_cookies import ChromeCookieDecryptor
from .config import get_config
from .errors import Mcp365Error
from .http import request

_OK, _WARN, _FAIL = "✅", "⚠️", "❌"


def _fmt_time(ts: float | None) -> str:
    if not ts:
        return "không rõ"
    return datetime.fromtimestamp(ts).strftime("%Y-%m-%d %H:%M")


def _remaining(ts: float | None) -> str:
    if not ts:
        return ""
    delta = ts - time.time()
    if delta <= 0:
        return " (ĐÃ HẾT HẠN)"
    hours = delta / 3600
    return f" (còn {hours:.1f} giờ)" if hours < 48 else f" (còn {hours / 24:.1f} ngày)"


def _probe_browser() -> dict[str, Any]:
    cfg = get_config().browser
    try:
        dbs = ChromeCookieDecryptor._candidate_cookie_dbs()
        ChromeCookieDecryptor.get_master_key()
        return {
            "status": _OK,
            "title": "Trình duyệt & Keyring",
            "detail": f"{cfg.name} · profile `{dbs[0].parent.name}` · master key giải mã được",
            "fix": "",
        }
    except Mcp365Error as exc:
        return {"status": _FAIL, "title": "Trình duyệt & Keyring", "detail": exc.message, "fix": exc.remediation}


def _probe_teams() -> dict[str, Any]:
    from teams.auth import TeamsAuthManager

    try:
        auth = TeamsAuthManager.get_auth()
    except Mcp365Error as exc:
        return {"status": _FAIL, "title": "Teams (skypetoken)", "detail": exc.message, "fix": exc.remediation}

    identity = auth["identity"]
    try:
        request(
            f"{auth['base_url']}/users/ME/conversations?view=msnp24Equivalent&pageSize=1",
            headers={"Authentication": f"skypetoken={auth['token']}", "Accept": "application/json"},
            context="kiểm tra Teams Chat Service",
            max_retries=0,
        )
    except Mcp365Error as exc:
        return {"status": _FAIL, "title": "Teams (skypetoken)", "detail": exc.message, "fix": exc.remediation}

    mt = "có" if auth.get("middle_tier_token") else "KHÔNG (tool Lịch sẽ không dùng được)"
    return {
        "status": _OK,
        "title": "Teams (skypetoken)",
        "detail": (
            f"{identity.display_name or identity.upn} · region `{auth['region']}` · "
            f"hết hạn {_fmt_time(auth['exp'])}{_remaining(auth['exp'])} · middle-tier token: {mt}"
        ),
        "fix": "",
    }


def _probe_sharepoint_cookies() -> dict[str, Any]:
    cfg = get_config().sharepoint
    try:
        cookies = ChromeCookieDecryptor.get_cookies_for_domain("sharepoint.com", ["rtFa", "FedAuth"])
    except Mcp365Error as exc:
        return {"status": _FAIL, "title": "SharePoint (cookie phiên)", "detail": exc.message, "fix": exc.remediation}

    missing = [n for n in ("rtFa", "FedAuth") if not cookies.get(n)]
    if missing:
        return {
            "status": _FAIL,
            "title": "SharePoint (cookie phiên)",
            "detail": f"Thiếu cookie: {', '.join(missing)}",
            "fix": f"Mở https://{cfg.hostname} trong Chrome và đăng nhập (nhớ tick 'Stay signed in').",
        }

    expiry = ChromeCookieDecryptor.get_cookie_expiry("sharepoint.com", "FedAuth")
    meta = f"FedAuth hết hạn {_fmt_time(expiry)}{_remaining(expiry)}"
    headers = {
        "Cookie": f"rtFa={cookies['rtFa']}; FedAuth={cookies['FedAuth']}",
        "User-Agent": get_config().http.user_agent,
    }

    # Page render and REST API are separate permissions. A cookie that renders
    # pages but fails /_api is the classic non-persistent-session symptom.
    page_ok = False
    try:
        request(cfg.site_url, headers={**headers, "Accept": "*/*"}, context="mở trang SharePoint", max_retries=0)
        page_ok = True
    except Mcp365Error:
        page_ok = False

    try:
        request(
            f"{cfg.site_url}/_api/web/title",
            headers={**headers, "Accept": "application/json;odata=verbose"},
            context="gọi SharePoint REST API",
            max_retries=0,
        )
        return {"status": _OK, "title": "SharePoint (cookie phiên)", "detail": f"REST API OK · {meta}", "fix": ""}
    except Mcp365Error as exc:
        hint = (
            " Trang web tải được (HTTP 200) nhưng REST API bị từ chối — dấu hiệu điển hình của phiên "
            "đăng nhập không bền."
            if page_ok
            else ""
        )
        return {
            "status": _FAIL,
            "title": "SharePoint (cookie phiên)",
            "detail": f"{exc.message}{hint} · {meta}",
            "fix": exc.remediation,
        }


def _probe_graph() -> dict[str, Any]:
    from sharepoint.client import SharePointClient

    client = SharePointClient()
    try:
        token = client.get_token()
    except Mcp365Error as exc:
        return {"status": _FAIL, "title": "Microsoft Graph (Azure CLI)", "detail": exc.message, "fix": exc.remediation}

    scopes = ""
    try:
        from teams.auth import decode_jwt_claims

        claims = decode_jwt_claims(token)
        scopes = claims.get("scp", "")
    except Exception:
        scopes = ""

    try:
        request(
            "https://graph.microsoft.com/v1.0/me",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            context="kiểm tra Microsoft Graph",
            max_retries=0,
        )
    except Mcp365Error as exc:
        return {
            "status": _FAIL,
            "title": "Microsoft Graph (Azure CLI)",
            "detail": exc.message,
            "fix": exc.remediation,
        }

    needed = {"Files.ReadWrite.All", "Sites.ReadWrite.All", "Files.Read.All", "Sites.Read.All"}
    have = set(scopes.split())
    if not (needed & have):
        return {
            "status": _WARN,
            "title": "Microsoft Graph (Azure CLI)",
            "detail": (
                "Token hợp lệ nhưng KHÔNG có scope Files.*/Sites.* — upload, replace và duyệt cây thư mục "
                f"có thể bị 403.\nScope hiện có: {scopes or 'không đọc được'}"
            ),
            "fix": "Chạy: az login --scope https://graph.microsoft.com/.default (hoặc nhờ admin cấp quyền Files.ReadWrite.All).",
        }
    return {"status": _OK, "title": "Microsoft Graph (Azure CLI)", "detail": "Token hợp lệ, có scope file.", "fix": ""}


def _probe_mail() -> dict[str, Any]:
    from outlook.auth import MailAuthManager

    auth = MailAuthManager()
    try:
        token = auth.get_token()
    except Mcp365Error as exc:
        return {
            "status": _WARN,
            "title": "Outlook mail (delegated Graph)",
            "detail": exc.message,
            "fix": "Gọi `start_mail_login`, nhập mã thiết bị, rồi gọi `check_mail_login`.",
        }

    try:
        request(
            "https://graph.microsoft.com/v1.0/me/mailFolders/inbox?$select=id",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            context="kiểm tra Outlook mail",
            max_retries=0,
        )
    except Mcp365Error as exc:
        return {
            "status": _FAIL,
            "title": "Outlook mail (delegated Graph)",
            "detail": exc.message,
            "fix": exc.remediation,
        }
    return {
        "status": _OK,
        "title": "Outlook mail (delegated Graph)",
        "detail": "Mail.Read + Mail.Send hoạt động cho mailbox của người đang đăng nhập.",
        "fix": "",
    }


def _probe_az_cli() -> dict[str, Any]:
    try:
        res = subprocess.run(["az", "version", "-o", "tsv"], capture_output=True, text=True, timeout=30)
        version = (res.stdout or "").split("\t")[0].strip() or "không rõ"
        return {"status": _OK, "title": "Azure CLI", "detail": f"phiên bản {version}", "fix": ""}
    except FileNotFoundError:
        return {
            "status": _WARN,
            "title": "Azure CLI",
            "detail": "Chưa cài 'az' — các tool ghi lên SharePoint sẽ không dùng được.",
            "fix": "Cài Azure CLI rồi chạy `az login`.",
        }
    except Exception as exc:
        return {"status": _WARN, "title": "Azure CLI", "detail": str(exc)[:160], "fix": ""}


def run_health_check() -> str:
    cfg = get_config()
    probes = [_probe_browser(), _probe_teams(), _probe_sharepoint_cookies(), _probe_mail(), _probe_az_cli(), _probe_graph()]

    failures = [p for p in probes if p["status"] == _FAIL]
    warnings = [p for p in probes if p["status"] == _WARN]
    if failures:
        headline = f"{_FAIL} {len(failures)} kênh đang hỏng"
    elif warnings:
        headline = f"{_WARN} Hoạt động, có {len(warnings)} cảnh báo"
    else:
        headline = f"{_OK} Tất cả các kênh đều hoạt động"

    out = [
        "# 🩺 Kiểm tra kết nối Microsoft 365",
        f"**{headline}** · {datetime.now().strftime('%Y-%m-%d %H:%M')}\n",
        "| | Kênh | Trạng thái |",
        "| --- | --- | --- |",
    ]
    for probe in probes:
        detail = probe["detail"].replace("\n", " ")
        out.append(f"| {probe['status']} | **{probe['title']}** | {detail} |")

    actions = [p for p in probes if p["fix"]]
    if actions:
        out.append("\n## 🔧 Việc cần làm")
        for idx, probe in enumerate(actions, 1):
            out.append(f"{idx}. **{probe['title']}** — {probe['fix']}")

    out.append("\n## ⚙️ Cấu hình đang dùng")
    out.append(f"- SharePoint site: `{cfg.sharepoint.site_url}`")
    out.append(f"- Trình duyệt: `{cfg.browser.name}` · profile `{cfg.browser.profile}`")
    out.append(f"- Outlook tenant: `{cfg.mail.tenant_id}` · delegated Mail.Read + Mail.Send")
    out.append(f"- Timeout: {cfg.http.timeout:.0f}s · retry: {cfg.http.max_retries} · workers: {cfg.http.max_workers}")
    out.append("\n> Đổi cấu hình qua `~/.config/mcp-auto-365-ms/config.toml` hoặc biến môi trường `MCP365_*`.")
    return "\n".join(out)
