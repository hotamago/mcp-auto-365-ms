"""Outlook mail client backed by delegated Microsoft Graph APIs."""

from __future__ import annotations

import html
import json
import re
import urllib.parse
from datetime import UTC, datetime
from typing import Any

from common.errors import ConfigError
from common.http import request, request_json

from .auth import MailAuthManager

_GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
_FOLDER_ALIASES = {
    "inbox": "inbox",
    "sent": "sentitems",
    "sentitems": "sentitems",
    "drafts": "drafts",
    "deleted": "deleteditems",
    "deleteditems": "deleteditems",
    "archive": "archive",
    "junk": "junkemail",
    "junkemail": "junkemail",
}


def clean_mail_body(content: str, content_type: str = "text") -> str:
    """Turn Graph message content into readable plain text."""
    if not content:
        return ""
    if content_type.casefold() != "html":
        return content.strip()
    text = re.sub(r"<br\s*/?>", "\n", content, flags=re.IGNORECASE)
    text = re.sub(r"</(p|div|li|tr|h[1-6])>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = html.unescape(text)
    text = re.sub(r"[ \t]+\n", "\n", text)
    return re.sub(r"\n{3,}", "\n\n", text).strip()


def _address(entry: dict[str, Any] | None) -> str:
    info = (entry or {}).get("emailAddress") or {}
    name, address = info.get("name", ""), info.get("address", "")
    return f"{name} <{address}>" if name and address and name != address else (address or name)


def _addresses(entries: list[dict[str, Any]] | None) -> list[str]:
    return [value for value in (_address(entry) for entry in entries or []) if value]


def _iso_since(value: str) -> str:
    raw = value.strip()
    if not raw:
        return ""
    try:
        parsed = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError as exc:
        raise ConfigError(
            f"Mốc thời gian mail không hợp lệ: `{value}`.",
            "Dùng YYYY-MM-DD hoặc ISO 8601, ví dụ `2026-09-20T08:00:00+07:00`.",
        ) from exc
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=datetime.now().astimezone().tzinfo)
    return parsed.astimezone(UTC).isoformat().replace("+00:00", "Z")


class OutlookMailClient:
    def __init__(self, auth: MailAuthManager | None = None) -> None:
        self.auth = auth or MailAuthManager()

    def _headers(self, prefer_text: bool = False) -> dict[str, str]:
        headers = {
            "Authorization": f"Bearer {self.auth.get_token()}",
            "Accept": "application/json",
        }
        if prefer_text:
            headers["Prefer"] = 'outlook.body-content-type="text"'
        return headers

    def _graph_json(
        self,
        path: str,
        *,
        method: str = "GET",
        body: dict[str, Any] | None = None,
        prefer_text: bool = False,
        context: str,
    ) -> dict[str, Any]:
        data = json.dumps(body).encode("utf-8") if body is not None else None
        headers = self._headers(prefer_text=prefer_text)
        if data is not None:
            headers["Content-Type"] = "application/json"
        return request_json(
            path if path.startswith("http") else f"{_GRAPH_ROOT}{path}",
            headers=headers,
            method=method,
            data=data,
            context=context,
        )

    @staticmethod
    def _format_message(raw: dict[str, Any], include_body: bool = False) -> dict[str, Any]:
        body = raw.get("body") or {}
        result = {
            "id": raw.get("id", ""),
            "conversation_id": raw.get("conversationId", ""),
            "subject": raw.get("subject") or "(không có tiêu đề)",
            "from": _address(raw.get("from")),
            "to": _addresses(raw.get("toRecipients")),
            "cc": _addresses(raw.get("ccRecipients")),
            "received": raw.get("receivedDateTime", ""),
            "sent": raw.get("sentDateTime", ""),
            "is_read": bool(raw.get("isRead")),
            "has_attachments": bool(raw.get("hasAttachments")),
            "preview": (raw.get("bodyPreview") or "").strip(),
            "web_link": raw.get("webLink", ""),
        }
        if include_body:
            result["body"] = clean_mail_body(body.get("content", ""), body.get("contentType", "text"))
        return result

    def list_messages(
        self,
        *,
        folder: str = "inbox",
        limit: int = 20,
        unread_only: bool = False,
        since: str = "",
        query: str = "",
    ) -> list[dict[str, Any]]:
        """List recent or searched messages from one mailbox folder."""
        folder_key = folder.strip().casefold().replace(" ", "") or "inbox"
        folder_id = _FOLDER_ALIASES.get(folder_key, folder.strip())
        if not folder_id or any(ch in folder_id for ch in "/?#"):
            raise ConfigError(
                f"Thư mục mail không hợp lệ: `{folder}`.",
                "Dùng inbox, sent, drafts, deleted, archive, junk hoặc Graph folder ID.",
            )

        limit = max(1, min(int(limit), 50))
        if query and (unread_only or since):
            raise ConfigError(
                "`query` không kết hợp với `unread_only` hoặc `since` vì Microsoft Graph không bảo đảm "
                "$search + $filter cho mail.",
                "Gọi tìm kiếm bằng `query` riêng, hoặc dùng `unread_only`/`since` mà không có `query`.",
            )
        since_iso = _iso_since(since) if since else ""
        params: dict[str, str] = {
            "$select": (
                "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,"
                "sentDateTime,isRead,hasAttachments,bodyPreview,webLink"
            ),
            "$top": str(limit),
        }
        if query:
            params["$search"] = f'"{query.strip()}"'
        else:
            filters = []
            if since_iso:
                # Ordered properties must appear first in Graph's $filter.
                filters.append(f"receivedDateTime ge {since_iso}")
            if unread_only:
                filters.append("isRead eq false")
            if filters:
                params["$filter"] = " and ".join(filters)
                if since_iso:
                    params["$orderby"] = "receivedDateTime desc"
            else:
                params["$orderby"] = "receivedDateTime desc"

        encoded_folder = urllib.parse.quote(folder_id, safe="")
        path = f"/me/mailFolders/{encoded_folder}/messages?{urllib.parse.urlencode(params)}"
        data = self._graph_json(path, prefer_text=True, context=f"đọc thư mục mail {folder}")
        return [self._format_message(item) for item in data.get("value", [])]

    def get_message(self, message_id: str) -> dict[str, Any]:
        """Read a single message, including its full body."""
        if not message_id.strip():
            raise ConfigError("Thiếu message_id của email.", "Lấy ID từ kết quả `list_emails` rồi gọi lại.")
        select = (
            "id,conversationId,subject,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,"
            "isRead,hasAttachments,bodyPreview,body,webLink"
        )
        encoded_id = urllib.parse.quote(message_id.strip(), safe="")
        path = f"/me/messages/{encoded_id}?{urllib.parse.urlencode({'$select': select})}"
        raw = self._graph_json(path, prefer_text=True, context="đọc nội dung email")
        return self._format_message(raw, include_body=True)

    def send_message(
        self,
        *,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
    ) -> dict[str, Any]:
        """Submit one plain-text message and save it to Sent Items."""
        recipients = [address.strip() for address in to if address.strip()]
        if not recipients:
            raise ConfigError("Email phải có ít nhất một người nhận.", "Truyền `to` dưới dạng danh sách địa chỉ email.")
        if not subject.strip():
            raise ConfigError("Email phải có tiêu đề.", "Điền `subject` trước khi xin người dùng duyệt.")

        def graph_recipients(addresses: list[str] | None) -> list[dict[str, dict[str, str]]]:
            return [{"emailAddress": {"address": address.strip()}} for address in addresses or [] if address.strip()]

        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": graph_recipients(recipients),
                "ccRecipients": graph_recipients(cc),
                "bccRecipients": graph_recipients(bcc),
            },
            "saveToSentItems": True,
        }
        data = json.dumps(payload).encode("utf-8")
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        status, _response_body, _response_headers = request(
            f"{_GRAPH_ROOT}/me/sendMail",
            headers=headers,
            method="POST",
            data=data,
            context="gửi email Outlook",
        )
        return {
            "status": status,
            "to": recipients,
            "cc": [address.strip() for address in cc or [] if address.strip()],
            "bcc": [address.strip() for address in bcc or [] if address.strip()],
            "subject": subject,
        }
