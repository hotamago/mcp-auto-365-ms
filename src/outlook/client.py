"""Outlook mail client using the signed-in browser's Outlook Web session."""

from __future__ import annotations

import html
import json
import re
import urllib.parse
from datetime import UTC, datetime
from typing import Any

from common.config import get_config
from common.errors import AuthExpiredError, ConfigError
from common.http import request, request_json

from .auth import MailAuthManager

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
    """Turn an Outlook message body into readable plain text."""
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


def _field(data: dict[str, Any], pascal: str, camel: str, default: Any = "") -> Any:
    return data.get(pascal, data.get(camel, default))


def _address(entry: dict[str, Any] | None) -> str:
    info = _field(entry or {}, "EmailAddress", "emailAddress", {}) or {}
    name = _field(info, "Name", "name")
    address = _field(info, "Address", "address")
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

    def _headers(self, prefer_text: bool = False, force_refresh: bool = False) -> dict[str, str]:
        headers = {
            "Authorization": f"Bearer {self.auth.get_token(force_refresh=force_refresh)}",
            "Accept": "application/json",
        }
        if prefer_text:
            headers["Prefer"] = 'outlook.body-content-type="text"'
        return headers

    def _outlook_json(self, path: str, *, prefer_text: bool = False, context: str) -> dict[str, Any]:
        api_root = get_config().mail.api_root.rstrip("/")
        url = path if path.startswith("http") else f"{api_root}{path}"
        try:
            return request_json(url, headers=self._headers(prefer_text=prefer_text), context=context)
        except AuthExpiredError:
            # Reads are idempotent. Mint once from the current Chrome session in
            # case the in-memory Outlook token expired early.
            return request_json(
                url,
                headers=self._headers(prefer_text=prefer_text, force_refresh=True),
                context=context,
            )

    @staticmethod
    def _format_message(raw: dict[str, Any], include_body: bool = False) -> dict[str, Any]:
        body = _field(raw, "Body", "body", {}) or {}
        result = {
            "id": _field(raw, "Id", "id"),
            "conversation_id": _field(raw, "ConversationId", "conversationId"),
            "subject": _field(raw, "Subject", "subject") or "(không có tiêu đề)",
            "from": _address(_field(raw, "From", "from", {})),
            "to": _addresses(_field(raw, "ToRecipients", "toRecipients", [])),
            "cc": _addresses(_field(raw, "CcRecipients", "ccRecipients", [])),
            "received": _field(raw, "ReceivedDateTime", "receivedDateTime"),
            "sent": _field(raw, "SentDateTime", "sentDateTime"),
            "is_read": bool(_field(raw, "IsRead", "isRead", False)),
            "has_attachments": bool(_field(raw, "HasAttachments", "hasAttachments", False)),
            "preview": str(_field(raw, "BodyPreview", "bodyPreview") or "").strip(),
            "web_link": _field(raw, "WebLink", "webLink"),
        }
        if include_body:
            result["body"] = clean_mail_body(
                _field(body, "Content", "content"),
                _field(body, "ContentType", "contentType", "text"),
            )
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
                "Dùng inbox, sent, drafts, deleted, archive, junk hoặc Outlook folder ID.",
            )

        limit = max(1, min(int(limit), 50))
        if query and (unread_only or since):
            raise ConfigError(
                "`query` không kết hợp với `unread_only` hoặc `since` vì Outlook không bảo đảm "
                "$search + $filter cho mail.",
                "Gọi tìm kiếm bằng `query` riêng, hoặc dùng `unread_only`/`since` mà không có `query`.",
            )
        since_iso = _iso_since(since) if since else ""
        params: dict[str, str] = {
            "$select": (
                "Id,ConversationId,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,"
                "SentDateTime,IsRead,HasAttachments,BodyPreview,WebLink"
            ),
            "$top": str(limit),
        }
        if query:
            params["$search"] = f'"{query.strip()}"'
        else:
            filters = []
            if since_iso:
                filters.append(f"ReceivedDateTime ge {since_iso}")
            if unread_only:
                filters.append("IsRead eq false")
            if filters:
                params["$filter"] = " and ".join(filters)
                if since_iso:
                    params["$orderby"] = "ReceivedDateTime desc"
            else:
                params["$orderby"] = "ReceivedDateTime desc"

        encoded_folder = urllib.parse.quote(folder_id, safe="")
        path = f"/me/mailfolders/{encoded_folder}/messages?{urllib.parse.urlencode(params)}"
        data = self._outlook_json(path, prefer_text=True, context=f"đọc thư mục mail {folder}")
        return [self._format_message(item) for item in data.get("value", [])]

    def get_message(self, message_id: str) -> dict[str, Any]:
        """Read a single message, including its full body."""
        if not message_id.strip():
            raise ConfigError("Thiếu message_id của email.", "Lấy ID từ kết quả `list_emails` rồi gọi lại.")
        select = (
            "Id,ConversationId,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,SentDateTime,"
            "IsRead,HasAttachments,BodyPreview,Body,WebLink"
        )
        encoded_id = urllib.parse.quote(message_id.strip(), safe="")
        path = f"/me/messages/{encoded_id}?{urllib.parse.urlencode({'$select': select})}"
        raw = self._outlook_json(path, prefer_text=True, context="đọc nội dung email")
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
        """Submit one plain-text message through the current Outlook Web session."""
        recipients = [address.strip() for address in to if address.strip()]
        if not recipients:
            raise ConfigError("Email phải có ít nhất một người nhận.", "Truyền `to` dưới dạng danh sách địa chỉ email.")
        if not subject.strip():
            raise ConfigError("Email phải có tiêu đề.", "Điền `subject` trước khi xin người dùng duyệt.")

        def outlook_recipients(addresses: list[str] | None) -> list[dict[str, dict[str, str]]]:
            return [{"EmailAddress": {"Address": address.strip()}} for address in addresses or [] if address.strip()]

        payload = {
            "Message": {
                "Subject": subject,
                "Body": {"ContentType": "Text", "Content": body},
                "ToRecipients": outlook_recipients(recipients),
                "CcRecipients": outlook_recipients(cc),
                "BccRecipients": outlook_recipients(bcc),
            },
            "SaveToSentItems": True,
        }
        data = json.dumps(payload).encode("utf-8")
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        status, _response_body, _response_headers = request(
            f"{get_config().mail.api_root.rstrip('/')}/me/sendmail",
            headers=headers,
            method="POST",
            data=data,
            context="gửi email qua phiên Outlook Web",
        )
        return {
            "status": status,
            "to": recipients,
            "cc": [address.strip() for address in cc or [] if address.strip()],
            "bcc": [address.strip() for address in bcc or [] if address.strip()],
            "subject": subject,
        }
