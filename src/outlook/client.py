"""Outlook mail client using the signed-in browser's Outlook Web session.

Mọi lệnh đi qua Outlook REST v2.0 (``mail.api_root``) bằng token Outlook Web lấy từ phiên
Chrome (``aud`` = ``https://outlook.office.com``). Token này có sẵn ``Mail.ReadWrite`` và
``Mail.Send`` (kiểm 28/09), nên soạn nháp, gắn file, trả lời/chuyển tiếp và tải file đính
kèm đều dùng được mà không cần Graph consent.

Gửi mail có file hoặc trả lời/chuyển tiếp đi theo một đường duy nhất: tạo bản nháp →
gắn file → ``/send``. Hỏng giữa chừng thì xoá bản nháp, không để mail dở dang.
"""

from __future__ import annotations

import base64
import html
import json
import mimetypes
import re
import urllib.parse
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from common.config import get_config
from common.errors import AuthExpiredError, ConfigError, Mcp365Error
from common.http import decode_json, request, request_json, request_to_file
from sharepoint.client import human_size

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

#: File nhỏ hơn mức này gắn thẳng bằng ``ContentBytes`` trong một request. Outlook REST giới
#: hạn thân request ~4 MB, base64 phình 4/3 nên phải *nhỏ hơn hẳn* 3 MiB.
DIRECT_ATTACH_LIMIT = 3 * 1024 * 1024
#: Giới hạn của upload session cho một file (kiểm 28/09: 149 MiB được, 151 MiB bị từ chối
#: ``ErrorAttachmentSizeShouldNotBeMoreThanMaximumSize``). Tổng các file gắn trực tiếp cũng
#: giữ dưới mức này; giới hạn gửi thật của tổ chức (MaxSendSize) REST không cho đọc.
SESSION_ATTACH_LIMIT = 150 * 1024 * 1024
#: Mỗi mảnh upload session: bội số 320 KiB và không quá 4 MB.
UPLOAD_CHUNK = 12 * 320 * 1024
#: Thư mục OneDrive cho file gửi dạng link, như Outlook tự làm với "tệp đính kèm đám mây".
MAIL_FILES_FOLDER = "Attachments"

ATTACH_MODES = ("auto", "attach", "link")
LINK_SCOPES = ("recipients", "organization")
BODY_FORMATS = ("text", "markdown")
REPLY_MODES = ("reply", "reply_all", "forward")

_PREFIX = {"reply": "RE: ", "reply_all": "RE: ", "forward": "FW: "}
_PREFIX_RE = {"reply": r"^(re|tl|trả lời)\s*:", "reply_all": r"^(re|tl|trả lời)\s*:", "forward": r"^(fw|fwd)\s*:"}
_CREATE_ACTION = {"reply": "createreply", "reply_all": "createreplyall", "forward": "createforward"}


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


# ------------------------------------------------------------------ nội dung

_MD_LINK = re.compile(r"\[([^\]\n]+)\]\(((?:https?://|mailto:)[^)\s]+)\)")
_MD_BOLD = re.compile(r"\*\*(?=\S)(.+?)(?<=\S)\*\*")
_MD_ITALIC = re.compile(r"(?<![*\w])\*(?=[^\s*])(.+?)(?<=[^\s*])\*(?![*\w])")
_MD_BULLET = re.compile(r"^\s*[-*+]\s+(.*)$")
_MD_NUMBER = re.compile(r"^\s*\d+[.)]\s+(.*)$")
_HTML_STYLE = "font-family:Calibri,Arial,sans-serif;font-size:11pt"


def _md_inline(text: str) -> str:
    """Escape trước, định dạng sau: văn bản người dùng không bao giờ thành thẻ HTML."""
    out = []
    parts = text.split("`")
    for i, part in enumerate(parts):
        # Phần lẻ nằm giữa hai dấu ` là code; thiếu dấu đóng thì giữ nguyên dấu `.
        if i % 2 == 1 and i < len(parts) - 1:
            out.append(f"<code>{html.escape(part)}</code>")
            continue
        piece = html.escape(part if i % 2 == 0 else "`" + part)
        piece = _MD_LINK.sub(lambda m: f'<a href="{m.group(2)}">{m.group(1)}</a>', piece)
        piece = _MD_BOLD.sub(r"<strong>\1</strong>", piece)
        piece = _MD_ITALIC.sub(r"<em>\1</em>", piece)
        out.append(piece)
    return "".join(out)


def markdown_to_html(text: str) -> str:
    """Markdown đơn giản → HTML: **đậm**, *nghiêng*, `code`, [link](url), gạch đầu dòng, xuống dòng.

    Link chỉ nhận http/https/mailto. Dòng trống tách đoạn, xuống dòng trong đoạn thành ``<br>``.
    """
    blocks: list[str] = []
    para: list[str] = []
    items: list[str] = []
    kind = ""

    def flush() -> None:
        nonlocal para, items, kind
        if para:
            blocks.append("<p>" + "<br>".join(para) + "</p>")
        if items:
            blocks.append(f"<{kind}>" + "".join(f"<li>{it}</li>" for it in items) + f"</{kind}>")
        para, items, kind = [], [], ""

    for line in text.replace("\r\n", "\n").split("\n"):
        bullet, number = _MD_BULLET.match(line), _MD_NUMBER.match(line)
        if bullet or number:
            want = "ul" if bullet else "ol"
            if para or (kind and kind != want):
                flush()
            kind = want
            items.append(_md_inline((bullet or number).group(1)))
        elif not line.strip():
            flush()
        else:
            if items:
                flush()
            para.append(_md_inline(line))
    flush()
    return f'<div style="{_HTML_STYLE}">' + "".join(blocks) + "</div>"


def text_to_html(text: str) -> str:
    """Văn bản thường → HTML giữ nguyên chữ và xuống dòng (không diễn giải ký tự nào)."""
    escaped = html.escape(text.replace("\r\n", "\n")).replace("\n", "<br>")
    return f'<div style="{_HTML_STYLE}">{escaped}</div>'


def render_body_html(body: str, body_format: str) -> str:
    if body_format not in BODY_FORMATS:
        raise ConfigError(
            f"body_format phải là 'text' hoặc 'markdown', không phải '{body_format}'.",
            "Dùng 'text' (mặc định, gửi nguyên văn) hoặc 'markdown' (**đậm**, *nghiêng*, [link](url), - gạch đầu dòng).",
        )
    return markdown_to_html(body) if body_format == "markdown" else text_to_html(body)


def _links_html(files: list[dict[str, Any]]) -> str:
    rows = "".join(
        f'<li><a href="{html.escape(f["share_link"], quote=True)}">{html.escape(f["name"])}</a>'
        f" ({html.escape(human_size(f['size']))})</li>"
        for f in files
    )
    return f'<div style="{_HTML_STYLE}"><p>📎 Tệp đính kèm (link OneDrive):</p><ul>{rows}</ul></div>'


def _insert_into_html(existing: str, ours: str) -> str:
    """Chèn phần mình viết lên đầu thân mail (ngay sau ``<body…>`` nếu có), giữ nguyên phần trích dẫn."""
    match = re.search(r"<body[^>]*>", existing, flags=re.IGNORECASE)
    if match:
        return existing[: match.end()] + ours + existing[match.end():]
    return ours + existing


# ---------------------------------------------------------------- file kèm


def plan_attachments(paths: list[str] | None, attach_mode: str = "auto") -> list[dict[str, Any]]:
    """Quyết định cách gắn từng file chỉ từ kích thước cục bộ (chưa gọi mạng).

    - ``direct``: nhỏ hơn 3 MiB, gắn thẳng.
    - ``session``: 3 MiB tới 150 MiB, gắn qua upload session của Outlook.
    - ``link``: ``attach_mode="link"``, hoặc (với ``auto``) file quá 150 MiB / tổng file gắn
      vượt 150 MiB → tải lên OneDrive của người gửi rồi chèn link vào thân mail.
    """
    if attach_mode not in ATTACH_MODES:
        raise ConfigError(
            f"attach_mode phải là 'auto', 'attach' hoặc 'link', không phải '{attach_mode}'.",
            "'auto' (mặc định) gắn trực tiếp khi được, quá giới hạn thì tự chuyển sang link OneDrive.",
        )
    planned: list[dict[str, Any]] = []
    total = 0
    for raw in paths or []:
        if not str(raw).strip():
            continue
        local = Path(str(raw).strip()).expanduser().resolve()
        if not local.is_file():
            raise ConfigError(f"Không tìm thấy file đính kèm: {raw}", "Kiểm tra lại đường dẫn file cục bộ.")
        size = local.stat().st_size
        if attach_mode == "link":
            method = "link"
        elif size > SESSION_ATTACH_LIMIT or total + size > SESSION_ATTACH_LIMIT:
            if attach_mode == "attach":
                raise ConfigError(
                    f"`{local.name}` ({human_size(size)}) làm tổng file gắn trực tiếp vượt giới hạn Outlook "
                    f"{human_size(SESSION_ATTACH_LIMIT)}.",
                    "Dùng attach_mode='auto' hoặc 'link' để gửi file lớn dạng link OneDrive.",
                )
            method = "link"
        else:
            method = "direct" if size < DIRECT_ATTACH_LIMIT else "session"
            total += size
        planned.append({"path": str(local), "name": local.name, "size": size, "method": method})
    return planned


def describe_attachment(item: dict[str, Any], link_scope: str = "recipients") -> str:
    """Một dòng cho bản nháp: tên, kích thước, cách gắn."""
    how = {
        "direct": "gắn trực tiếp",
        "session": "gắn trực tiếp (upload session, file ≥ 3 MB)",
    }.get(item["method"])
    if not how:
        who = (
            "chỉ người nhận To/CC/BCC mở được (cấp quyền từng người, không gửi mail mời)"
            if link_scope == "recipients"
            else "mọi người trong tổ chức có link đều mở được"
        )
        how = f"link: tải lên OneDrive của bạn › `{MAIL_FILES_FOLDER}`, {who}"
    return f"`{item['name']}` · {human_size(item['size'])} · {how}"


# ---------------------------------------------------------------- tiện ích


def _field(data: dict[str, Any], pascal: str, camel: str, default: Any = "") -> Any:
    return data.get(pascal, data.get(camel, default))


def _raw_address(entry: dict[str, Any] | None) -> str:
    info = _field(entry or {}, "EmailAddress", "emailAddress", {}) or {}
    return str(_field(info, "Address", "address") or "").strip()


def _address(entry: dict[str, Any] | None) -> str:
    info = _field(entry or {}, "EmailAddress", "emailAddress", {}) or {}
    name = _field(info, "Name", "name")
    address = _field(info, "Address", "address")
    return f"{name} <{address}>" if name and address and name != address else (address or name)


def _addresses(entries: list[dict[str, Any]] | None) -> list[str]:
    return [value for value in (_address(entry) for entry in entries or []) if value]


def _raw_addresses(entries: list[dict[str, Any]] | None) -> list[str]:
    return [value for value in (_raw_address(entry) for entry in entries or []) if value]


def clean_addresses(addresses: list[str] | None) -> list[str]:
    """Bỏ khoảng trắng, bỏ trùng (không phân biệt hoa thường), giữ thứ tự."""
    seen: set[str] = set()
    out: list[str] = []
    for address in addresses or []:
        value = str(address).strip()
        if value and value.casefold() not in seen:
            seen.add(value.casefold())
            out.append(value)
    return out


def _outlook_recipients(addresses: list[str] | None) -> list[dict[str, dict[str, str]]]:
    return [{"EmailAddress": {"Address": address}} for address in clean_addresses(addresses)]


def _safe_file_name(name: str, fallback: str) -> str:
    cleaned = re.sub(r'[\\/:*?"<>|\x00-\x1f]', "_", name or "").strip(" .")
    return cleaned or fallback


def _free_path(folder: Path, name: str) -> Path:
    """``folder/name`` hoặc ``name (1).ext``… nếu đã có file: không bao giờ ghi đè."""
    target = folder / name
    stem, suffix = target.stem, target.suffix
    n = 1
    while target.exists():
        target = folder / f"{stem} ({n}){suffix}"
        n += 1
    return target


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


def _odata_string(value: str) -> str:
    return value.replace("'", "''")


class OutlookMailClient:
    def __init__(self, auth: MailAuthManager | None = None, sharepoint: Any = None) -> None:
        self.auth = auth or MailAuthManager()
        self._sharepoint = sharepoint
        self._me: dict[str, str] | None = None

    # ------------------------------------------------------------ HTTP

    def _headers(self, prefer_text: bool = False, force_refresh: bool = False) -> dict[str, str]:
        headers = {
            "Authorization": f"Bearer {self.auth.get_token(force_refresh=force_refresh)}",
            "Accept": "application/json",
        }
        if prefer_text:
            headers["Prefer"] = 'outlook.body-content-type="text"'
        return headers

    @staticmethod
    def _url(path: str) -> str:
        return path if path.startswith("http") else f"{get_config().mail.api_root.rstrip('/')}{path}"

    @staticmethod
    def _mid(message_id: str) -> str:
        if not message_id.strip():
            raise ConfigError("Thiếu message_id của email.", "Lấy ID từ kết quả `list_emails` rồi gọi lại.")
        return urllib.parse.quote(message_id.strip(), safe="")

    def _outlook_json(self, path: str, *, prefer_text: bool = False, context: str) -> dict[str, Any]:
        url = self._url(path)
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

    def _write(self, path: str, *, method: str = "POST", body: dict[str, Any] | None = None, context: str) -> dict:
        """Một lệnh ghi (không tự thử lại: POST có thể đã tới máy chủ)."""
        headers = self._headers()
        data = None
        if body is not None:
            headers["Content-Type"] = "application/json"
            data = json.dumps(body).encode("utf-8")
        url = self._url(path)
        _status, raw, _hdrs = request(url, headers=headers, method=method, data=data, context=context)
        return decode_json(raw, url)

    def _sp(self):
        if self._sharepoint is None:
            from sharepoint.client import SharePointClient

            self._sharepoint = SharePointClient()
        return self._sharepoint

    def me(self) -> dict[str, str]:
        """Địa chỉ mailbox của chính mình (``GET /me``). UPN trong token có thể khác địa chỉ mail."""
        if self._me is None:
            data = self._outlook_json("/me", context="đọc địa chỉ mailbox của bạn")
            self._me = {
                "address": str(_field(data, "EmailAddress", "emailAddress") or ""),
                "name": str(_field(data, "DisplayName", "displayName") or ""),
                "alias": str(_field(data, "Alias", "alias") or ""),
            }
        return self._me

    def _my_addresses(self) -> set[str]:
        me = self.me()
        found = {me["address"].casefold()} if me["address"] else set()
        try:
            from teams.auth import decode_jwt_claims

            claims = decode_jwt_claims(self.auth.get_token())
            for key in ("upn", "unique_name", "preferred_username"):
                if claims.get(key):
                    found.add(str(claims[key]).casefold())
        except Exception:  # noqa: BLE001 - chỉ để nhận diện thêm, thiếu cũng không sao
            pass
        return found

    # ------------------------------------------------------------ đọc

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
            "importance": str(_field(raw, "Importance", "importance") or "Normal"),
            "is_flagged": _field(_field(raw, "Flag", "flag", {}) or {}, "FlagStatus", "flagStatus") == "Flagged",
            "preview": str(_field(raw, "BodyPreview", "bodyPreview") or "").strip(),
            "web_link": _field(raw, "WebLink", "webLink"),
        }
        if include_body:
            result["body"] = clean_mail_body(
                _field(body, "Content", "content"),
                _field(body, "ContentType", "contentType", "text"),
            )
            result["raw_subject"] = _field(raw, "Subject", "subject") or ""
            result["from_address"] = _raw_address(_field(raw, "From", "from", {}))
            result["to_addresses"] = _raw_addresses(_field(raw, "ToRecipients", "toRecipients", []))
            result["cc_addresses"] = _raw_addresses(_field(raw, "CcRecipients", "ccRecipients", []))
            result["reply_to_addresses"] = _raw_addresses(_field(raw, "ReplyTo", "replyTo", []))
        return result

    def _folder_id(self, folder: str) -> str:
        """Tên thường dùng (inbox, sent…), Outlook folder ID, hoặc tên hiển thị của thư mục.

        Tên hiển thị được tìm ở thư mục cấp đầu và thư mục con của Inbox (nơi rule hay đặt).
        """
        raw = folder.strip()
        key = raw.casefold().replace(" ", "") or "inbox"
        if key in _FOLDER_ALIASES:
            return _FOLDER_ALIASES[key]
        if any(ch in raw for ch in "/?#"):
            raise ConfigError(
                f"Thư mục mail không hợp lệ: `{folder}`.",
                "Dùng inbox, sent, drafts, deleted, archive, junk, tên thư mục hoặc Outlook folder ID.",
            )
        if len(raw) >= 60 and " " not in raw:
            return raw  # Outlook folder ID
        params = urllib.parse.urlencode(
            {"$filter": f"DisplayName eq '{_odata_string(raw)}'", "$select": "Id,DisplayName", "$top": "5"}
        )
        for base in ("/me/mailfolders", "/me/mailfolders/inbox/childfolders"):
            data = self._outlook_json(f"{base}?{params}", context=f"tìm thư mục mail '{raw}'")
            for item in data.get("value", []):
                if str(_field(item, "DisplayName", "displayName")).casefold() == raw.casefold():
                    return str(_field(item, "Id", "id"))
        raise ConfigError(
            f"Không tìm thấy thư mục mail `{folder}`.",
            "Kiểm tra tên thư mục (cấp đầu hoặc thư mục con của Inbox), hoặc dùng inbox, sent, drafts, "
            "deleted, archive, junk.",
        )

    def list_messages(
        self,
        *,
        folder: str = "inbox",
        limit: int = 20,
        unread_only: bool = False,
        since: str = "",
        query: str = "",
        sender: str = "",
        has_attachments: bool = False,
    ) -> list[dict[str, Any]]:
        """List recent or searched messages from one mailbox folder.

        ``sender`` là địa chỉ email thì lọc chính xác (kết hợp được với các bộ lọc khác); là tên
        thì chuyển thành tìm kiếm ``from:"tên"``.
        """
        sender = sender.strip()
        sender_by_search = bool(sender) and "@" not in sender
        search_text = query.strip()
        if sender_by_search:
            # Outlook nhận KQL trong một chuỗi ngoặc kép, ngoặc kép bên trong phải thoát: "from:\"tên\" từ khoá"
            search_text = 'from:\\"' + sender.replace('"', "") + '\\"' + (f" {search_text}" if search_text else "")
        filtering = unread_only or bool(since) or has_attachments or (bool(sender) and not sender_by_search)
        if search_text and filtering:
            raise ConfigError(
                "Tìm kiếm (`query` hoặc `sender` là tên) không kết hợp với `unread_only`, `since`, "
                "`has_attachments` hay `sender` là email vì Outlook không bảo đảm $search + $filter cho mail.",
                "Gọi tìm kiếm riêng, hoặc lọc bằng `sender` là địa chỉ email kèm `unread_only`/`since`/"
                "`has_attachments` mà không có `query`.",
            )
        folder_id = self._folder_id(folder)
        limit = max(1, min(int(limit), 50))
        since_iso = _iso_since(since) if since else ""
        params: dict[str, str] = {
            "$select": (
                "Id,ConversationId,Subject,From,ToRecipients,CcRecipients,ReceivedDateTime,"
                "SentDateTime,IsRead,HasAttachments,Importance,Flag,BodyPreview,WebLink"
            ),
            "$top": str(limit),
        }
        if search_text:
            params["$search"] = f'"{search_text}"'
        else:
            filters = []
            if since_iso:
                filters.append(f"ReceivedDateTime ge {since_iso}")
            if unread_only:
                filters.append("IsRead eq false")
            if has_attachments:
                filters.append("HasAttachments eq true")
            if sender:
                filters.append(f"From/EmailAddress/Address eq '{_odata_string(sender)}'")
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
        select = (
            "Id,ConversationId,Subject,From,ToRecipients,CcRecipients,ReplyTo,ReceivedDateTime,SentDateTime,"
            "IsRead,HasAttachments,Importance,Flag,BodyPreview,Body,WebLink"
        )
        path = f"/me/messages/{self._mid(message_id)}?{urllib.parse.urlencode({'$select': select})}"
        raw = self._outlook_json(path, prefer_text=True, context="đọc nội dung email")
        return self._format_message(raw, include_body=True)

    def list_attachments(self, message_id: str) -> list[dict[str, Any]]:
        """Tên, kích thước, loại của file đính kèm (không kéo nội dung file về)."""
        params = urllib.parse.urlencode({"$select": "Id,Name,Size,ContentType,IsInline"})
        data = self._outlook_json(
            f"/me/messages/{self._mid(message_id)}/attachments?{params}", context="liệt kê file đính kèm"
        )
        out = []
        for item in data.get("value", []):
            odata_type = str(item.get("@odata.type") or "")
            kind = "item" if "ItemAttachment" in odata_type else "reference" if "Reference" in odata_type else "file"
            out.append(
                {
                    "id": str(_field(item, "Id", "id")),
                    "name": str(_field(item, "Name", "name") or "attachment"),
                    "size": int(_field(item, "Size", "size", 0) or 0),
                    "content_type": str(_field(item, "ContentType", "contentType") or ""),
                    "is_inline": bool(_field(item, "IsInline", "isInline", False)),
                    "kind": kind,
                }
            )
        return out

    def download_attachments(
        self, message_id: str, target_dir: str = "", name_filter: str = "", include_inline: bool = False
    ) -> dict[str, Any]:
        """Tải file đính kèm của một mail về ``target_dir`` (không ghi đè file có sẵn).

        File thường: ``/$value`` (luồng byte), không được thì ``ContentBytes``. Mail đính kèm:
        lưu ``.eml``. Tệp đám mây (link OneDrive/SharePoint): tải qua kênh SharePoint.
        """
        dest = Path(target_dir or "downloads").expanduser().resolve()
        wanted = name_filter.strip().casefold()
        items = [
            a for a in self.list_attachments(message_id)
            if (include_inline or not a["is_inline"]) and (not wanted or wanted in a["name"].casefold())
        ]
        saved: list[dict[str, Any]] = []
        failures: list[str] = []
        mid = self._mid(message_id)
        for item in items:
            aid = urllib.parse.quote(item["id"], safe="")
            try:
                dest.mkdir(parents=True, exist_ok=True)
                if item["kind"] == "reference":
                    data = self._outlook_json(f"/me/messages/{mid}/attachments/{aid}", context="đọc link tệp đám mây")
                    source = str(_field(data, "SourceUrl", "sourceUrl") or "")
                    if not source:
                        raise Mcp365Error(f"Tệp đám mây `{item['name']}` không có link nguồn.")
                    report = self._sp().download_link(source, target_dir=str(dest))
                    saved.append({"name": item["name"], "path": "", "size": item["size"], "report": report})
                    continue
                name = _safe_file_name(item["name"], f"attachment-{len(saved) + 1}")
                if item["kind"] == "item" and not name.lower().endswith(".eml"):
                    name += ".eml"
                target = _free_path(dest, name)
                size = self._save_attachment(mid, aid, item, target)
                saved.append({"name": item["name"], "path": str(target), "size": size})
            except Mcp365Error as exc:
                failures.append(f"`{item['name']}`: {exc.message}")
        return {"target_dir": str(dest), "matched": len(items), "saved": saved, "failures": failures}

    def _save_attachment(self, mid: str, aid: str, item: dict[str, Any], target: Path) -> int:
        url = self._url(f"/me/messages/{mid}/attachments/{aid}/$value")
        try:
            return request_to_file(url, target, headers=self._headers(), context=f"tải `{item['name']}`")
        except Mcp365Error as first:
            if item["kind"] == "item":
                raise
            data = self._outlook_json(f"/me/messages/{mid}/attachments/{aid}", context=f"tải `{item['name']}`")
            content = _field(data, "ContentBytes", "contentBytes")
            if not content:
                raise first
            blob = base64.b64decode(content)
            target.write_bytes(blob)
            return len(blob)

    # ------------------------------------------------------------ gửi

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
        recipients = clean_addresses(to)
        if not recipients:
            raise ConfigError("Email phải có ít nhất một người nhận.", "Truyền `to` dưới dạng danh sách địa chỉ email.")
        if not subject.strip():
            raise ConfigError("Email phải có tiêu đề.", "Điền `subject` trước khi xin người dùng duyệt.")

        payload = {
            "Message": {
                "Subject": subject,
                "Body": {"ContentType": "Text", "Content": body},
                "ToRecipients": _outlook_recipients(recipients),
                "CcRecipients": _outlook_recipients(cc),
                "BccRecipients": _outlook_recipients(bcc),
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
            "cc": clean_addresses(cc),
            "bcc": clean_addresses(bcc),
            "subject": subject,
        }

    def prepare_reply(
        self,
        message_id: str,
        mode: str = "reply",
        to: list[str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
    ) -> dict[str, Any]:
        """Đọc mail gốc và tính đúng To/CC/BCC + tiêu đề sẽ gửi (chưa ghi gì).

        - ``reply``: tới người gửi (hoặc Reply-To của mail gốc).
        - ``reply_all``: người gửi + To gốc, CC gốc; bỏ chính mình.
        - ``forward``: chỉ những người trong ``to``/``cc``/``bcc``.
        ``to``/``cc``/``bcc`` truyền vào được thêm vào danh sách trên.
        """
        if mode not in REPLY_MODES:
            raise ConfigError(
                f"mode phải là 'reply', 'reply_all' hoặc 'forward', không phải '{mode}'.",
                "'reply' trả lời người gửi, 'reply_all' trả lời tất cả, 'forward' chuyển tiếp (cần `to`).",
            )
        original = self.get_message(message_id)
        mine = self._my_addresses() if mode == "reply_all" else set()

        def not_me(addresses: list[str]) -> list[str]:
            return [a for a in addresses if a.casefold() not in mine]

        sender = original["reply_to_addresses"] or ([original["from_address"]] if original["from_address"] else [])
        if mode == "reply":
            base_to, base_cc = sender, []
        elif mode == "reply_all":
            base_to = not_me(clean_addresses(sender + original["to_addresses"])) or sender
            base_cc = not_me(original["cc_addresses"])
        else:
            base_to, base_cc = [], []
        final_to = clean_addresses(base_to + (to or []))
        taken = {a.casefold() for a in final_to}
        final_cc = [a for a in clean_addresses(base_cc + (cc or [])) if a.casefold() not in taken]
        taken |= {a.casefold() for a in final_cc}
        final_bcc = [a for a in clean_addresses(bcc) if a.casefold() not in taken]
        if not final_to:
            raise ConfigError(
                "Chuyển tiếp cần ít nhất một người nhận." if mode == "forward" else "Mail gốc không có người gửi để trả lời.",
                "Truyền `to` là danh sách địa chỉ email.",
            )
        subject = original["raw_subject"]
        if not re.match(_PREFIX_RE[mode], subject.strip(), flags=re.IGNORECASE):
            subject = _PREFIX[mode] + subject
        return {"original": original, "mode": mode, "to": final_to, "cc": final_cc, "bcc": final_bcc, "subject": subject}

    def resolve_upns(self, addresses: list[str]) -> dict[str, str]:
        """Địa chỉ mail → UPN trong danh bạ công ty (SharePoint cấp quyền theo UPN, có thể khác mail)."""
        found: dict[str, str] = {}
        for address in clean_addresses(addresses):
            query = urllib.parse.quote(f'"{address}"')
            data = self._outlook_json(f"/me/people?$top=10&$search={query}", context=f"tra danh bạ '{address}'")
            for person in data.get("value", []):
                emails = {
                    str(e.get("Address") or "").casefold()
                    for key in ("ScoredEmailAddresses", "EmailAddresses")
                    for e in person.get(key, []) or []
                }
                upn = str(person.get("UserPrincipalName") or "")
                if upn and (address.casefold() in emails or upn.casefold() == address.casefold()):
                    found[address] = upn
                    break
        return found

    def _share_link_files(
        self, files: list[dict[str, Any]], recipients: list[str], link_scope: str
    ) -> list[dict[str, Any]]:
        """Tải file lên OneDrive › Attachments rồi tạo link cho đúng người nhận (hoặc cả tổ chức).

        Người nhận được tra UPN *trước khi* tải file lên; thiếu ai thì dừng, chưa tải gì.
        """
        if link_scope not in LINK_SCOPES:
            raise ConfigError(
                f"link_scope phải là 'recipients' hoặc 'organization', không phải '{link_scope}'.",
                "'recipients' (mặc định): chỉ người nhận mail mở được; 'organization': cả công ty có link.",
            )
        upns: list[str] = []
        if link_scope == "recipients":
            mine = self._my_addresses()
            others = [a for a in clean_addresses(recipients) if a.casefold() not in mine]
            found = self.resolve_upns(others)
            missing = [a for a in others if a not in found]
            if missing:
                raise Mcp365Error(
                    "Email CHƯA được gửi: không tìm thấy trong danh bạ công ty "
                    f"{', '.join(missing)} để cấp quyền file (người ngoài công ty không mở được link loại này).",
                    "Gắn file trực tiếp (attach_mode='attach'), hoặc dùng link_scope='organization' nếu người đó "
                    "trong công ty, hoặc bỏ file khỏi mail.",
                )
            upns = [found[a] for a in others]
        sp = self._sp()
        from sharepoint.client import view_url

        shared = []
        for item in files:
            info = sp.upload_unique(item["path"], MAIL_FILES_FOLDER)
            remote = f"{MAIL_FILES_FOLDER}/{info['name']}"
            try:
                if link_scope == "organization":
                    link = sp.create_org_link(info["drive_id"], remote, "view")
                elif upns:
                    link = sp.create_people_link(info["drive_id"], remote, upns, "view")
                else:  # chỉ gửi cho chính mình
                    link = view_url(info.get("fileUrl") or info.get("webUrl") or "")
            except Mcp365Error as exc:
                raise Mcp365Error(
                    f"Email CHƯA được gửi: `{info['name']}` đã lên OneDrive của bạn ({remote}) nhưng không cấp "
                    f"quyền được.\n{exc.message}",
                    exc.remediation or "Xoá file đó trên OneDrive nếu không cần, rồi thử lại.",
                ) from exc
            shared.append({**item, "name": info["name"], "share_link": link, "location": f"OneDrive › {remote}"})
        return shared

    def _add_attachment(self, draft_id: str, item: dict[str, Any]) -> None:
        path = Path(item["path"])
        content_type = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
        mid = self._mid(draft_id)
        if item["method"] == "direct":
            self._write(
                f"/me/messages/{mid}/attachments",
                body={
                    "@odata.type": "#Microsoft.OutlookServices.FileAttachment",
                    "Name": item["name"],
                    "ContentType": content_type,
                    "ContentBytes": base64.b64encode(path.read_bytes()).decode("ascii"),
                },
                context=f"gắn `{item['name']}`",
            )
            return
        session = self._write(
            f"/me/messages/{mid}/attachments/createuploadsession",
            body={"AttachmentItem": {"AttachmentType": "File", "Name": item["name"], "Size": item["size"],
                                     "ContentType": content_type}},
            context=f"mở upload session cho `{item['name']}`",
        )
        upload_url = str(_field(session, "UploadUrl", "uploadUrl") or "")
        if not upload_url:
            raise Mcp365Error(f"Outlook không trả UploadUrl để gắn `{item['name']}`.")
        size, sent = item["size"], 0
        with path.open("rb") as fh:
            while sent < size:
                blob = fh.read(UPLOAD_CHUNK)
                if not blob:
                    break
                self._put_chunk(upload_url, blob, sent, size, f"gắn `{item['name']}` (mảnh {sent // UPLOAD_CHUNK + 1})")
                sent += len(blob)

    @staticmethod
    def _put_chunk(upload_url: str, blob: bytes, start: int, size: int, context: str, attempts: int = 3) -> None:
        """PUT một mảnh; tự thử lại khi mạng lỗi.

        Outlook nhận mảnh rồi mới mất phản hồi thì lần gửi lại bị trả 400 ``InvalidStart``
        ("Fragment might already have been uploaded") - gặp thật 28/09. Session không có lệnh
        GET trạng thái (405), nên ``InvalidStart`` ở lần thử lại được hiểu là mảnh đã tới.
        """
        for attempt in range(attempts):
            try:
                # Không gửi Authorization: UploadUrl đã mang authtoken riêng.
                request(
                    upload_url,
                    headers={
                        "Content-Type": "application/octet-stream",
                        "Content-Range": f"bytes {start}-{start + len(blob) - 1}/{size}",
                    },
                    method="PUT",
                    data=blob,
                    kind="transfer",
                    max_retries=0,
                    context=context,
                )
                return
            except Mcp365Error as exc:
                if attempt and exc.http_status == 400 and "InvalidStart" in exc.message:
                    return
                retryable = exc.http_status in (None, 429, 500, 502, 503, 504)
                if not retryable or attempt == attempts - 1:
                    raise

    def _discard_draft(self, draft_id: str) -> str:
        try:
            self._write(f"/me/messages/{self._mid(draft_id)}", method="DELETE", context="xoá bản nháp dở")
            return "Bản nháp dở đã được xoá."
        except Mcp365Error as exc:
            return f"Không xoá được bản nháp dở (ID `{draft_id}`): {exc.message}. Xoá tay trong Drafts."

    def compose_and_send(
        self,
        *,
        to: list[str],
        subject: str,
        body: str,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        body_format: str = "text",
        attachments: list[dict[str, Any]] | None = None,
        link_scope: str = "recipients",
        reply_to_id: str = "",
        mode: str = "",
    ) -> dict[str, Any]:
        """Soạn nháp → gắn file → gửi. ``attachments`` là kết quả của :func:`plan_attachments`.

        ``reply_to_id`` + ``mode`` (reply/reply_all/forward): bản nháp tạo từ mail gốc nên giữ
        đúng thread và phần trích dẫn; To/CC/BCC và tiêu đề được đặt lại đúng như bản đã duyệt.
        Mail mới, văn bản thường, không file: đi đường ``/sendmail`` cũ.
        """
        to, cc, bcc = clean_addresses(to), clean_addresses(cc), clean_addresses(bcc)
        attachments = attachments or []
        if not to:
            raise ConfigError("Email phải có ít nhất một người nhận.", "Truyền `to` dưới dạng danh sách địa chỉ email.")
        if not subject.strip():
            raise ConfigError("Email phải có tiêu đề.", "Điền `subject` trước khi xin người dùng duyệt.")
        ours = render_body_html(body, body_format)
        if not reply_to_id and not attachments and body_format == "text":
            result = self.send_message(to=to, subject=subject, body=body, cc=cc, bcc=bcc)
            result.update({"attachments": [], "draft_id": "", "mode": "new"})
            return result

        linked = [a for a in attachments if a["method"] == "link"]
        attached = [a for a in attachments if a["method"] != "link"]
        shared = self._share_link_files(linked, to + cc + bcc, link_scope) if linked else []
        if shared:
            ours += _links_html(shared)

        fields = {
            "Subject": subject,
            "ToRecipients": _outlook_recipients(to),
            "CcRecipients": _outlook_recipients(cc),
            "BccRecipients": _outlook_recipients(bcc),
        }
        try:
            if reply_to_id:
                if mode not in REPLY_MODES:
                    raise ConfigError(f"mode phải là một trong {', '.join(REPLY_MODES)}.")
                draft = self._write(
                    f"/me/messages/{self._mid(reply_to_id)}/{_CREATE_ACTION[mode]}",
                    body={},
                    context="tạo bản nháp trả lời/chuyển tiếp",
                )
            else:
                draft = self._write(
                    "/me/messages",
                    body={**fields, "Body": {"ContentType": "HTML", "Content": ours}},
                    context="tạo bản nháp email",
                )
        except Mcp365Error as exc:
            uploaded = f"\nFile đã lên OneDrive: {', '.join(f['location'] for f in shared)}." if shared else ""
            raise Mcp365Error(f"Email CHƯA được gửi: {exc.message}{uploaded}", exc.remediation) from exc
        draft_id = str(_field(draft, "Id", "id") or "")
        if not draft_id:
            raise Mcp365Error("Outlook không trả ID bản nháp; email CHƯA được gửi.")
        try:
            if reply_to_id:
                existing = _field(draft, "Body", "body", {}) or {}
                content = str(_field(existing, "Content", "content") or "")
                if str(_field(existing, "ContentType", "contentType") or "").casefold() != "html":
                    content = text_to_html(content)
                self._write(
                    f"/me/messages/{self._mid(draft_id)}",
                    method="PATCH",
                    body={**fields, "Body": {"ContentType": "HTML", "Content": _insert_into_html(content, ours)}},
                    context="điền bản nháp trả lời/chuyển tiếp",
                )
            for item in attached:
                self._add_attachment(draft_id, item)
        except Mcp365Error as exc:
            note = self._discard_draft(draft_id)
            raise Mcp365Error(f"Email CHƯA được gửi: {exc.message}\n{note}", exc.remediation) from exc
        try:
            status, _raw, _hdrs = request(
                self._url(f"/me/messages/{self._mid(draft_id)}/send"),
                headers=self._headers(),
                method="POST",
                context="gửi email qua phiên Outlook Web",
            )
        except Mcp365Error as exc:
            raise Mcp365Error(
                f"Gửi email thất bại: {exc.message}\nBản nháp (ID `{draft_id}`) vẫn nằm trong Drafts.",
                "Xem Sent Items trước khi thử lại (lỗi mạng có thể đã gửi rồi). Mail quá giới hạn dung lượng "
                "của tổ chức thì gửi lại với attach_mode='link'.",
            ) from exc
        return {
            "status": status,
            "to": to,
            "cc": cc,
            "bcc": bcc,
            "subject": subject,
            "draft_id": draft_id,
            "conversation_id": str(_field(draft, "ConversationId", "conversationId") or ""),
            "mode": mode or "new",
            "attachments": attached + shared,
        }
