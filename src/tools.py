"""Tool, resource and prompt registrations.

Declared once and mounted onto any server instance. Previously the unified
server and the two standalone servers each re-declared the same tools, and the
copies had already drifted apart (one still advertised a hardcoded user name).
"""

from __future__ import annotations

import functools
import urllib.parse
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from mcp.server.mcpserver.exceptions import ToolError

from common import approval
from common.errors import Mcp365Error
from common.health import run_health_check
from sharepoint import docx_comments, sheets
from sharepoint.client import SharePointClient, human_size
from teams.client import TeamsClient

_sp_client: SharePointClient | None = None
_teams_client: TeamsClient | None = None


def sp() -> SharePointClient:
    """Lazily built SharePoint client.

    Nothing may be constructed at import time: the Teams client used to call the
    keyring in its constructor, so a locked keyring took down all 16 tools -
    including the SharePoint ones, which do not even need that credential.
    """
    global _sp_client
    if _sp_client is None:
        _sp_client = SharePointClient()
    return _sp_client


def teams() -> TeamsClient:
    global _teams_client
    if _teams_client is None:
        _teams_client = TeamsClient()
    return _teams_client


def _actionable(fn):
    """Re-raise our typed errors as ``ToolError`` so the text survives.

    The SDK passes a ``ToolError`` message through verbatim but replaces any
    other exception with a bare "Error executing tool <name>" - which would
    discard exactly the remediation the user needs.
    """

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        try:
            return fn(*args, **kwargs)
        except Mcp365Error as exc:
            raise ToolError(str(exc)) from exc

    return wrapper


class _ErrorAwareServer:
    """Proxy that applies :func:`_actionable` to every registered tool."""

    def __init__(self, mcp) -> None:
        self._mcp = mcp

    def tool(self, *args, **kwargs):
        decorator = self._mcp.tool(*args, **kwargs)

        def wrap(fn):
            return decorator(_actionable(fn))

        return wrap

    def __getattr__(self, name):
        return getattr(self._mcp, name)


def _errors_note(errors: list[str]) -> str:
    """Render partial failures instead of silently returning a short answer."""
    if not errors:
        return ""
    lines = "\n".join(f"> - {e}" for e in errors[:10])
    more = f"\n> - …và {len(errors) - 10} lỗi khác" if len(errors) > 10 else ""
    return f"\n\n> ⚠️ **Không quét được {len(errors)} cuộc trò chuyện** (kết quả bên trên chưa đầy đủ):\n{lines}{more}"


def _render_messages(messages: list[dict[str, Any]], bullet: bool = True) -> list[str]:
    out = []
    for msg in messages:
        links = msg.get("sharepoint_links") or []
        suffix = ("\n  *Links:* " + ", ".join(f"[Link]({link})" for link in links)) if links else ""
        # A file-only message has an empty body; without this it rendered as a
        # blank line and the attachment was invisible to whoever read the chat.
        files = msg.get("attachments") or []
        if files:
            suffix += "\n  📎 " + ", ".join(f"`{f['name']}`" for f in files)
        flag = " 🔔" if msg.get("mentions_me") else ""
        if bullet:
            out.append(f"- **[{msg['timestamp']}] {msg['sender']}{flag}:** {msg['content']}{suffix}")
        else:
            out.append(f"### [{msg['timestamp']}] {msg['sender']}{flag}\n{msg['content']}{suffix}\n")
    return out


# ============================================================ registration


def register_sharepoint_tools(mcp) -> None:
    mcp = _ErrorAwareServer(mcp)
    @mcp.tool()
    def search_sharepoint_files(query: str, max_results: int = 20, file_extension: str = "") -> str:
        """Search SharePoint/OneDrive documents by keyword, with optional file-type filter.

        Args:
            query: Keyword to search for (e.g. 'SYS2', 'CAN', 'Architecture').
            max_results: Maximum number of files to return.
            file_extension: Optional extension filter (e.g. 'docx', 'xlsx', 'pdf').
        """
        results = sp().search_files(query=query, max_results=max_results, file_extension=file_extension or None)
        if not results:
            return f"Không tìm thấy tài liệu nào khớp với '{query}'."
        out = [
            f"# Kết quả tìm kiếm SharePoint cho '{query}' ({len(results)} tài liệu)\n",
            "| Tài liệu | Kích thước | Sửa lần cuối | Tác giả | UniqueId |",
            "| --- | --- | --- | --- | --- |",
        ]
        for r in results:
            out.append(
                f"| **[{r['title']}]({r['path']})** | {human_size(r['size'])} | {r['modified'][:10]} "
                f"| {r['author'][:25]} | `{r['unique_id']}` |"
            )
        out.append("\n> Gọi `download_sharepoint_link(unique_id)` để tải bất kỳ file nào.")
        return "\n".join(out)

    @mcp.tool()
    def read_sharepoint_link(url: str, max_depth: int = 2) -> str:
        """Explore a SharePoint folder tree, or show a document's metadata and version history.

        Args:
            url: SharePoint folder or document URL.
            max_depth: How many folder levels to traverse.
        """
        return sp().read_link(url, max_depth=max_depth)

    @mcp.tool()
    def download_sharepoint_link(url_or_guid: str, target_dir: str = "") -> str:
        """Download original binary files or a whole folder from SharePoint, with no format conversion.

        Args:
            url_or_guid: SharePoint URL, sharing link, or document UniqueId (GUID).
            target_dir: Destination directory (defaults to the configured download dir).
        """
        return sp().download_link(url_or_guid, target_dir=target_dir)

    @mcp.tool()
    def upload_sharepoint_file(
        local_file_path: str,
        target_folder_url_or_path: str,
        is_user_confirm: approval.UserConfirm,
        target_file_name: str = "",
    ) -> str:
        """Upload a local file to SharePoint, creating any missing parent folders.

        Ask the user before calling with is_user_confirm=true.

        Args:
            local_file_path: Path to the local file.
            target_folder_url_or_path: Destination folder URL or site-relative path.
            is_user_confirm: Required. True only after the user approved this exact upload.
            target_file_name: Optional remote filename (defaults to the local name).
        """
        approval.require_confirm(
            is_user_confirm,
            "Tải file lên SharePoint",
            target_folder_url_or_path,
            f"`{local_file_path}` → `{target_file_name or Path(local_file_path).name}`",
        )
        res = sp().upload_file(local_file_path, target_folder_url_or_path, target_file_name or None)
        return (
            f"✓ Đã tải `{res['name']}` ({human_size(res['size'])}) lên SharePoint.\n"
            f"- **Thư mục:** `{res['folder']}`\n- **Item ID:** `{res['id']}`\n- **Web URL:** {res['webUrl']}"
        )

    @mcp.tool()
    def replace_sharepoint_file(local_file_path: str, file_url_or_guid: str, is_user_confirm: approval.UserConfirm) -> str:
        """Replace an existing SharePoint file in place, creating a new version and keeping its link and ID.

        Ask the user before calling with is_user_confirm=true.

        Args:
            local_file_path: Path to the updated local file.
            file_url_or_guid: SharePoint file URL, sharing link, or UniqueId.
            is_user_confirm: Required. True only after the user approved overwriting this file.
        """
        approval.require_confirm(
            is_user_confirm, "Ghi đè file trên SharePoint", file_url_or_guid, f"Thay nội dung bằng `{local_file_path}`"
        )
        return _render_replace(sp().replace_file(local_file_path, file_url_or_guid))

    def _render_replace(res: dict[str, Any]) -> str:
        return (
            f"✓ Đã thay thế `{res['name']}` ({human_size(res['size'])}) trên SharePoint.\n"
            f"- **Phiên bản mới:** `{res['version']}`\n- **Item ID:** `{res['id']}`\n"
            f"- **Thời điểm sửa:** {res['modified']}\n- **Web URL:** {res['webUrl']}"
        )

    @mcp.tool()
    def read_sharepoint_sheet(file_url_or_guid: str, sheet: str = "", max_rows: int = 60) -> str:
        """Read an Excel workbook on SharePoint/OneDrive: list its sheets, or show one as a table.

        Rows are numbered and columns lettered, so the output gives the exact A1
        addresses to pass to `update_sharepoint_sheet`.

        Args:
            file_url_or_guid: File URL (any site or OneDrive), sharing link, or UniqueId.
            sheet: Sheet to show. Empty lists every sheet with its size.
            max_rows: Maximum rows to render.
        """
        drive_id, item = sp().resolve_file(file_url_or_guid)
        data = sp().read_file_bytes(drive_id, item)
        return sheets.render_sheet(data, sheet=sheet, max_rows=max_rows, name=item.get("name", ""))

    @mcp.tool()
    def add_sharepoint_docx_comments(
        file_url_or_guid: str,
        comments: list[dict[str, str]],
        is_user_confirm: approval.UserConfirm,
        author: str = "",
    ) -> str:
        """Add review comments to a Word document on SharePoint, anchored to its text.

        Each comment is attached to the first paragraph containing its `anchor`
        phrase, exactly like a comment added in Word. Call first with
        is_user_confirm=false to get where each comment will land, show that to the
        user, and call again with true only after they approve. The upload uses
        `If-Match: <eTag>`, so if the document changed since it was read - or is open
        in a co-authoring session - nothing is written.

        Args:
            file_url_or_guid: Document URL (any site or OneDrive), sharing link, or UniqueId.
            comments: List of {"anchor": short verbatim phrase from the document, "text": comment}.
            is_user_confirm: Required. True only after the user approved these exact comments.
            author: Comment author shown in Word. Defaults to the signed-in Teams user.
        """
        drive_id, item = sp().resolve_file(file_url_or_guid)
        etag = item.get("eTag", "")
        original = sp().read_file_bytes(drive_id, item)
        if not author:
            try:
                author = teams().identity.display_name
            except Mcp365Error:
                author = ""
        new_bytes, report = docx_comments.add_comments(original, comments, author or "Reviewer")
        approval.require_confirm(
            is_user_confirm,
            "Thêm comment vào tài liệu Word trên SharePoint",
            f"{item.get('name')} (tác giả: {author or 'Reviewer'})",
            docx_comments.render_report(report),
        )
        res = sp().put_file_bytes(drive_id, item["id"], new_bytes, if_match=etag)
        return (
            f"✓ Đã thêm {len(report)} comment vào `{item.get('name')}` (phiên bản mới).\n"
            f"- **Web URL:** {res.get('webUrl', item.get('webUrl', ''))}"
        )

    @mcp.tool()
    def update_sharepoint_sheet(
        file_url_or_guid: str,
        sheet: str,
        cells: dict[str, str],
        is_user_confirm: approval.UserConfirm,
        copy_sheet_from: str = "",
    ) -> str:
        """Edit cells of an Excel file on SharePoint without clobbering concurrent edits.

        Call first with is_user_confirm=false to get the exact change list, show it to
        the user, and only call again with true once they approve. The upload is sent
        with `If-Match: <eTag>`: if anyone saved the file since it was read - or an open
        co-authoring session locks it - nothing is written and the tool says so. Charts
        and images are dropped by the round trip; tables, styles, merged cells and
        comments survive.

        Args:
            file_url_or_guid: File URL (any site or OneDrive), sharing link, or UniqueId.
            sheet: Target sheet. Created if missing.
            cells: A1 address → value, e.g. {"D3": "No", "E3": "Thiếu link catalog S1–S3"}.
            is_user_confirm: Required. True only after the user approved this exact change list.
            copy_sheet_from: When `sheet` is missing, clone this sheet (rows + formatting) first.
        """
        drive_id, item = sp().resolve_file(file_url_or_guid)
        etag = item.get("eTag", "")
        original = sp().read_file_bytes(drive_id, item)
        new_bytes, changes = sheets.apply_cells(original, sheet, cells, copy_sheet_from, name=item.get("name", ""))
        approval.require_confirm(
            is_user_confirm, "Sửa ô Excel trên SharePoint", f"{item.get('name')} › {sheet}", sheets.render_changes(changes, sheet)
        )
        res = sp().put_file_bytes(drive_id, item["id"], new_bytes, if_match=etag)
        return (
            f"✓ Đã ghi {len(changes)} thay đổi vào `{item.get('name')}` › `{sheet}` (phiên bản mới).\n"
            f"- **Web URL:** {res.get('webUrl', item.get('webUrl', ''))}"
        )

    @mcp.tool()
    def compare_sharepoint_versions(file_a: str, file_b: str = "", version_a: str = "", version_b: str = "") -> str:
        """Diff two SharePoint document versions, or a local file against a SharePoint document.

        Args:
            file_a: Local path or SharePoint URL/GUID.
            file_b: Optional second file to compare against file_a.
            version_a: (When file_b is omitted) earlier version label, e.g. '1.0'.
            version_b: (When file_b is omitted) later version label, e.g. '2.0' or 'latest'.
        """
        if file_b:
            return sp().compare_documents(file_a, file_b)
        return sp().compare_versions(file_a, version_a=version_a, version_b=version_b)

    @mcp.tool()
    def sync_folder_to_sharepoint(
        local_dir: str, target_folder: str, is_user_confirm: approval.UserConfirm, dry_run: bool = True
    ) -> str:
        """Upload local files that are new or newer than their SharePoint copy.

        One-directional and non-destructive: never deletes anything remotely.
        Defaults to a dry run so the plan can be reviewed first; a real upload
        (dry_run=false) needs the user's approval of that plan.

        Args:
            local_dir: Local directory to sync from.
            target_folder: SharePoint destination folder URL or relative path.
            is_user_confirm: Required. For dry_run=false, true only after the user approved the plan.
            dry_run: When True (default) only reports what would be uploaded.
        """
        if not dry_run:
            approval.require_confirm(
                is_user_confirm, "Đồng bộ thư mục lên SharePoint", target_folder, f"Tải các file mới/đổi từ `{local_dir}`"
            )
        return sp().sync_folder_up(local_dir, target_folder, dry_run=dry_run)

    @mcp.tool()
    def download_meeting_recordings(target_dir: str = "", limit: int = 3, query: str = "Recording") -> str:
        """Find and download Teams meeting recordings stored in SharePoint/OneDrive.

        Args:
            target_dir: Destination directory.
            limit: Maximum number of recordings to download.
            query: Search term used to locate recordings.
        """
        return sp().download_meeting_recordings(target_dir=target_dir, limit=limit, query=query)


def register_teams_tools(mcp) -> None:
    mcp = _ErrorAwareServer(mcp)
    @mcp.tool()
    def list_teams_chats(limit: int = 30, filter_keyword: str = "", chat_type: str = "") -> str:
        """List recent Teams group chats, 1:1 chats, meeting chats and channels.

        The keyword is matched without diacritics, so ``nam son`` finds
        ``Nguyễn Phan Nam Sơn``, and it is applied across every known
        conversation before ``limit`` truncates the result.

        Args:
            limit: Maximum number of conversations to return.
            filter_keyword: Optional keyword filter on chat name or last message.
            chat_type: Optional type filter: DirectChat, GroupChat, Channel or MeetingChat.
        """
        chats = teams().list_conversations(page_size=limit, filter_keyword=filter_keyword, chat_type=chat_type)
        if not chats:
            return "Không tìm thấy cuộc trò chuyện nào khớp."
        out = [f"# Cuộc trò chuyện Microsoft Teams ({len(chats)})\n", "| Loại | Tên | Người gửi cuối | Hoạt động | Chat ID |", "| --- | --- | --- | --- | --- |"]
        for c in chats:
            when = c["last_activity"][:19].replace("T", " ") if c["last_activity"] else "N/A"
            out.append(f"| {c['type']} | **{c['name']}** | {c['last_sender']} | {when} | `{c['id']}` |")
        return "\n".join(out)

    @mcp.tool()
    def read_teams_chat(
        chat_name_or_id: str,
        limit: int = 30,
        since: str = "",
        only_mentions: bool = False,
        output_file: str = "",
    ) -> str:
        """Read message history from a Teams chat, channel or 1:1 conversation.

        Args:
            chat_name_or_id: Chat name (partial match works), or thread ID.
            limit: Number of recent messages to fetch.
            since: Optional time filter: 'today', 'yesterday', '6h', '3d' or 'YYYY-MM-DD' (local time).
            only_mentions: Only return messages that mention you.
            output_file: Optional path to also save the transcript as Markdown.
        """
        res = teams().get_messages(chat_name_or_id, limit=limit, since=since or None, only_mentions=only_mentions)
        if not res["messages"]:
            return f"Không có tin nhắn nào trong '{res['conversation_name']}' khớp điều kiện."
        out = [
            f"# Chat: {res['conversation_name']}",
            f"- **Loại:** {res['conversation_type']}",
            f"- **Thread ID:** `{res['conversation_id']}`",
            f"- **Số tin nhắn:** {len(res['messages'])}\n",
            "---",
        ]
        out.extend(_render_messages(res["messages"], bullet=False))
        rendered = "\n".join(out)
        if output_file:
            path = Path(output_file).expanduser().resolve()
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(rendered, encoding="utf-8")
            return f"✓ Đã lưu {len(res['messages'])} tin nhắn vào `{output_file}`.\n\n{rendered}"
        return rendered

    @mcp.tool()
    def get_recent_team_messages(
        hours: int = 48, max_chats: int = 8, limit_per_chat: int = 8, filter_keyword: str = ""
    ) -> str:
        """Fetch new messages across all active chats and channels in one parallel call.

        Args:
            hours: How many hours back to look.
            max_chats: Maximum conversations to scan.
            limit_per_chat: Maximum messages per conversation.
            filter_keyword: Optional filter on chat name or content.
        """
        res = teams().get_recent_feed(
            hours=hours, max_chats=max_chats, limit_per_chat=limit_per_chat, filter_keyword=filter_keyword
        )
        if not res["feed"]:
            return f"Không có tin nhắn mới trong {hours} giờ qua.{_errors_note(res['errors'])}"
        out = [f"# Hoạt động Teams gần đây ({hours} giờ qua — {len(res['feed'])}/{res['scanned']} chat có tin mới)\n"]
        for item in res["feed"]:
            out.append(f"## 💬 {item['chat_name']} *({item['chat_type']})*")
            out.append(f"- **Chat ID:** `{item['chat_id']}`\n")
            out.extend(_render_messages(item["messages"]))
            out.append("\n---")
        return "\n".join(out) + _errors_note(res["errors"])

    @mcp.tool()
    def get_my_mentions(hours: int = 72, limit: int = 20, context_before: int = 2, context_after: int = 2) -> str:
        """Find messages that mention you, across group chats, channels and 1:1 chats.

        Matching uses the authoritative mention payload Teams attaches to each
        message (your user MRI), so it works regardless of how your display name
        is rendered. Returns surrounding messages for context.

        Args:
            hours: How many hours back to look.
            limit: Maximum mentions to return.
            context_before: Messages to include before each mention (0-10).
            context_after: Messages to include after each mention (0-10).
        """
        res = teams().get_user_mentions(
            hours=hours,
            limit=limit,
            context_before=max(0, min(context_before, 10)),
            context_after=max(0, min(context_after, 10)),
        )
        if not res["mentions"]:
            return f"Không có tin nhắn nào nhắc tới bạn trong {hours} giờ qua.{_errors_note(res['errors'])}"
        out = [f"# Tin nhắn nhắc tới bạn ({len(res['mentions'])} trong {hours} giờ qua)\n"]
        for m in res["mentions"]:
            out.append(f"### 📍 [{m['chat_name']}] — **{m['sender']}** tag bạn ({m['timestamp']}) · *{m['mention_reason']}*")
            out.append(f"- **Message ID:** `{m['message_id']}` · **Chat ID:** `{m['chat_id']}`")
            ctx = m.get("context") or []
            if ctx:
                out.append("\n**Bối cảnh thảo luận:**")
                for c in ctx:
                    when = c["timestamp"][11:19] if len(c["timestamp"]) >= 19 else c["timestamp"]
                    if c["is_mention"]:
                        out.append(f"👉 **[{when}] {c['sender']} (MENTION):**\n> {c['content']}")
                    else:
                        out.append(f"- *({c['offset']:+d}) [{when}] {c['sender']}:* {c['content']}")
                    if c.get("sharepoint_links"):
                        out.append("  *Links:* " + ", ".join(f"[Link]({link})" for link in c["sharepoint_links"]))
            else:
                out.append(f"> {m['content']}")
            out.append("\n---\n")
        return "\n".join(out) + _errors_note(res["errors"])

    @mcp.tool()
    def get_new_mentions_since(cursor: str = "", limit: int = 20) -> str:
        """Return only mentions newer than a cursor, for periodic polling.

        Pass the cursor returned by the previous call. An MCP stdio server cannot
        hold a background watch loop, so the caller drives the polling (for
        example from a scheduled task).

        Args:
            cursor: Timestamp from the previous call ('YYYY-MM-DD HH:MM:SS'); empty scans the last 24h.
            limit: Maximum mentions to return.
        """
        res = teams().get_new_mentions_since(cursor=cursor, limit=limit)
        header = f"**Cursor tiếp theo:** `{res['cursor']}`"
        if not res["mentions"]:
            return f"Không có mention mới kể từ `{cursor or '24 giờ trước'}`.\n\n{header}{_errors_note(res['errors'])}"
        out = [f"# {len(res['mentions'])} mention mới\n", header, ""]
        for m in res["mentions"]:
            out.append(f"### 📍 [{m['chat_name']}] — **{m['sender']}** ({m['timestamp']})")
            out.append(f"- **Message ID:** `{m['message_id']}` · **Chat ID:** `{m['chat_id']}`")
            out.append(f"> {m['content']}\n")
        return "\n".join(out) + _errors_note(res["errors"])

    @mcp.tool()
    def search_teams_chat_messages(keywords: list[str], limit: int = 20) -> str:
        """Search recent messages across chats and channels for any of several keywords.

        Args:
            keywords: Keywords or phrases to look for (e.g. ['DTC', 'S5', 'review']).
            limit: Maximum matching messages to return.
        """
        res = teams().search_messages(keywords=keywords, max_results=limit)
        shown = ", ".join(f"'{k}'" for k in (keywords if isinstance(keywords, list) else [keywords]))
        if not res["hits"]:
            return f"Không tìm thấy tin nhắn nào chứa: [{shown}].{_errors_note(res['errors'])}"
        out = [f"# Kết quả tìm kiếm [{shown}] ({len(res['hits'])} kết quả)\n"]
        for hit in res["hits"]:
            out.append(
                f"**[{hit['chat_name']}]** — *{hit['sender']}* ({hit['timestamp']}) "
                f"*(khớp: `{hit['matched_keyword']}`)*:"
            )
            out.append(f"> {hit['content'][:300]}")
            out.append(f"> *Message ID:* `{hit['id']}`\n")
        return "\n".join(out) + _errors_note(res["errors"])

    @mcp.tool()
    def send_teams_message(
        chat_name_or_id: str,
        message: str,
        is_user_confirm: approval.UserConfirm,
        reply_to_id: str = "",
        file_path: str = "",
    ) -> str:
        """Send a Teams message, optionally quoting another message or attaching a local file.

        Microsoft Teams is sensitive: ALWAYS ask the user first. Show them the exact
        message and the destination chat, wait for an explicit yes, and only then
        call with is_user_confirm=true. With false, nothing is sent and the draft is
        returned for you to show them.

        Args:
            chat_name_or_id: Chat name (partial match works) or thread ID.
            message: Message text; **bold**, *italic*, `code` and [links](url) are supported.
            is_user_confirm: Required. True only after the user approved this exact message to this chat.
            reply_to_id: Optional message ID to quote-reply to.
            file_path: Optional local file to upload to SharePoint and attach.
        """
        conv = teams().find_conversation(chat_name_or_id)
        detail = message + (f"\n\n_(đính kèm: {file_path})_" if file_path else "")
        approval.require_confirm(is_user_confirm, "Gửi tin nhắn Teams", f"{conv['name']} (`{conv['id']}`)", detail)
        return _render_send(
            teams().send_message(
                conversation_id_or_name=conv["id"],
                message=message,
                reply_to_id=reply_to_id or None,
                file_path=file_path or None,
            )
        )

    def _render_send(res: dict[str, Any]) -> str:
        extra = []
        if res.get("reply_to_id"):
            extra.append(f"- **Trả lời tin nhắn:** `{res['reply_to_id']}`")
        if res.get("attached_file"):
            extra.append(f"- **File đính kèm:** [{res['attached_file']['name']}]({res['attached_file']['webUrl']})")
        if res.get("message_id"):
            extra.append(f"- **Message ID:** `{res['message_id']}`")
        suffix = "\n" + "\n".join(extra) if extra else ""
        return f"✓ Đã gửi tin nhắn tới **{res['conversation_name']}** (`{res['conversation_id']}`):{suffix}\n\n> {res['message_sent']}"

    @mcp.tool()
    def reply_to_channel_thread(
        channel_name_or_id: str, parent_message_id: str, message: str, is_user_confirm: approval.UserConfirm
    ) -> str:
        """Reply inside an existing Teams channel thread instead of starting a new one.

        Channels address a thread as `<channel-id>;messageid=<root>`; a plain send
        always creates a new thread, which is why this is a separate tool.
        ALWAYS ask the user first and send only after an explicit yes.

        Args:
            channel_name_or_id: Channel name (e.g. '[Team] #General') or thread ID.
            parent_message_id: ID of the thread's root message.
            message: Reply text.
            is_user_confirm: Required. True only after the user approved this exact reply.
        """
        conv = teams().find_conversation(channel_name_or_id)
        approval.require_confirm(
            is_user_confirm, "Trả lời thread trong channel", f"{conv['name']} · thread `{parent_message_id}`", message
        )
        return _render_reply(teams().reply_to_channel_thread(conv["id"], parent_message_id, message))

    def _render_reply(res: dict[str, Any]) -> str:
        return (
            f"✓ Đã trả lời trong thread của **{res['conversation_name']}**.\n"
            f"- **Thread:** `{res['thread_id']}`\n- **Message ID:** `{res['message_id']}`\n\n> {res['message_sent']}"
        )

    @mcp.tool()
    def edit_teams_message(
        chat_name_or_id: str, message_id: str, new_message: str, is_user_confirm: approval.UserConfirm
    ) -> str:
        """Edit one of your own previously sent Teams messages.

        ALWAYS ask the user first and edit only after an explicit yes.

        Args:
            chat_name_or_id: Chat name or thread ID.
            message_id: ID of the message to edit.
            new_message: Replacement text.
            is_user_confirm: Required. True only after the user approved this exact new text.
        """
        conv = teams().find_conversation(chat_name_or_id)
        approval.require_confirm(is_user_confirm, "Sửa tin nhắn Teams", f"{conv['name']} · tin `{message_id}`", new_message)
        res = teams().edit_message(conv["id"], message_id=message_id, new_message=new_message)
        return f"✓ Đã sửa tin nhắn `{res['message_id']}` trong '{res['conversation_name']}':\n{res['new_message']}"

    @mcp.tool()
    def delete_teams_message(chat_name_or_id: str, message_id: str, is_user_confirm: approval.UserConfirm) -> str:
        """Delete (recall) one of your own previously sent Teams messages.

        ALWAYS ask the user first and delete only after an explicit yes.

        Args:
            chat_name_or_id: Chat name or thread ID.
            message_id: ID of the message to delete.
            is_user_confirm: Required. True only after the user approved deleting this message.
        """
        conv = teams().find_conversation(chat_name_or_id)
        approval.require_confirm(
            is_user_confirm, "Xoá tin nhắn Teams", f"{conv['name']} (`{conv['id']}`)", f"Xoá tin nhắn `{message_id}`"
        )
        res = teams().delete_message(conv["id"], message_id=message_id)
        return f"✓ Đã xoá tin nhắn `{res['message_id']}` khỏi '{res['conversation_name']}'."

    @mcp.tool()
    def download_chat_attachments(
        chat_name_or_id: str, target_dir: str = "", limit: int = 5, file_name: str = "", scan_messages: int = 50
    ) -> str:
        """Download files from a chat: paperclip attachments and SharePoint/OneDrive links.

        Attachments live in the message's ``properties.files`` (the HTML body of a
        file-only message is empty), so both sources are scanned, newest first.

        Args:
            chat_name_or_id: Chat name or thread ID.
            target_dir: Destination directory.
            limit: Maximum number of files to download.
            file_name: Optional case-insensitive substring to pick one file, e.g. "PSDK.zip".
            scan_messages: How many recent messages to scan.
        """
        res = teams().get_messages(chat_name_or_id, limit=scan_messages)
        wanted = file_name.lower().strip()
        links: list[str] = []
        for msg in reversed(res["messages"]):
            candidates = [(a["name"], a["url"]) for a in msg.get("attachments", [])]
            candidates += [(link.rsplit("/", 1)[-1], link) for link in msg.get("sharepoint_links", [])]
            for name, link in candidates:
                if wanted and wanted not in urllib.parse.unquote(name).lower():
                    continue
                if link not in links:
                    links.append(link)
        if not links:
            what = f"file khớp '{file_name}'" if wanted else "file đính kèm hay link SharePoint/OneDrive nào"
            return f"Không tìm thấy {what} trong {scan_messages} tin gần nhất của '{res['conversation_name']}'."

        reports, failures = [], []
        for link in links[:limit]:
            try:
                reports.append(sp().download_link(link, target_dir=target_dir))
            except Mcp365Error as exc:
                failures.append(f"- `{link[:70]}…`: {exc.message}")
        body = f"# Đã xử lý {len(reports)}/{min(len(links), limit)} tệp từ '{res['conversation_name']}'\n\n"
        body += "\n\n---\n\n".join(reports)
        if failures:
            body += "\n\n> ⚠️ **Thất bại:**\n" + "\n".join(f"> {f}" for f in failures)
        return body

    @mcp.tool()
    def get_calendar_today(days: int = 1) -> str:
        """List your Teams calendar meetings, with join links.

        Uses the Teams middle-tier session from Chrome, because the Azure CLI
        Graph token carries no Calendars.* scope.

        Args:
            days: How many days ahead to include (1 = today only).
        """
        start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0).astimezone()
        end = start + timedelta(days=max(1, days))
        events = teams().get_calendar_events(start, end)
        if not events:
            return f"Không có cuộc họp nào từ {start:%Y-%m-%d} đến {end:%Y-%m-%d}."
        out = [f"# 📅 Lịch họp ({start:%Y-%m-%d} → {end:%Y-%m-%d}) — {len(events)} cuộc họp\n"]
        for ev in events:
            out.append(f"### {ev['subject']}")
            out.append(f"- **Thời gian:** {ev['start']} → {ev['end']}")
            if ev.get("organizer"):
                out.append(f"- **Người tổ chức:** {ev['organizer']}")
            if ev.get("join_url"):
                out.append(f"- **Tham gia:** [Teams meeting]({ev['join_url']})")
            out.append("")
        return "\n".join(out)


def register_shared_tools(mcp) -> None:
    mcp = _ErrorAwareServer(mcp)
    @mcp.tool()
    def check_365_connection() -> str:
        """Diagnose every Microsoft 365 auth channel and report exactly what to fix.

        Checks the browser keyring, the Teams session, SharePoint session cookies,
        the Azure CLI and the Graph token (including its scopes), and prints the
        concrete remediation for anything that is broken. Run this first whenever
        another tool fails with a permission error.
        """
        return run_health_check()

    @mcp.tool()
    def extract_action_items(hours: int = 72, limit: int = 25) -> str:
        """Collect messages that look like assigned work, as structured raw material.

        Returns mentions plus request-shaped messages with their chat, sender,
        timestamp and message ID, so they can be triaged into a task list. This
        tool does no summarising of its own - it gathers the evidence.

        Args:
            hours: How many hours back to scan.
            limit: Maximum items to return.
        """
        client = teams()
        mention_res = client.get_user_mentions(hours=hours, limit=limit, context_before=1, context_after=1)
        cue_words = ["giúp", "nhờ", "check", "review", "deadline", "gấp", "urgent", "cần", "hạn", "todo", "task", "fix"]
        keyword_res = client.search_messages(keywords=cue_words, max_results=limit)

        seen: set[str] = set()
        items: list[dict[str, Any]] = []
        for m in mention_res["mentions"]:
            key = str(m["message_id"])
            if key not in seen:
                seen.add(key)
                items.append({**m, "source": "mention"})
        for hit in keyword_res["hits"]:
            key = str(hit["id"])
            if key not in seen:
                seen.add(key)
                items.append(
                    {
                        "chat_name": hit["chat_name"],
                        "chat_id": hit["chat_id"],
                        "message_id": hit["id"],
                        "sender": hit["sender"],
                        "timestamp": hit["timestamp"],
                        "content": hit["content"],
                        "source": f"keyword:{hit['matched_keyword']}",
                        "context": [],
                    }
                )

        items.sort(key=lambda x: x["timestamp"], reverse=True)
        items = items[:limit]
        if not items:
            return f"Không tìm thấy đầu việc tiềm năng nào trong {hours} giờ qua."

        out = [
            f"# 📋 Nguyên liệu đầu việc ({len(items)} mục trong {hours} giờ qua)",
            "*Dữ liệu thô đã gom sẵn — hãy tự phân loại thành danh sách việc cần làm.*\n",
        ]
        for idx, item in enumerate(items, 1):
            out.append(f"## {idx}. [{item['chat_name']}] — {item['sender']} ({item['timestamp']})")
            out.append(f"- **Nguồn:** `{item['source']}` · **Message ID:** `{item['message_id']}` · **Chat ID:** `{item['chat_id']}`")
            out.append(f"> {item['content'][:500]}")
            for ctx in item.get("context") or []:
                if not ctx["is_mention"]:
                    out.append(f"  - *({ctx['offset']:+d}) {ctx['sender']}:* {ctx['content'][:160]}")
            out.append("")
        errs = mention_res["errors"] + keyword_res["errors"]
        return "\n".join(out) + _errors_note(errs)

    @mcp.tool()
    def get_daily_briefing(hours: int = 24) -> str:
        """Morning briefing: mentions, active discussions, calendar and recent documents.

        Args:
            hours: How many hours back to synthesise.
        """
        client = teams()
        sections = [
            f"# ☀️ Tổng hợp công việc Microsoft 365 ({datetime.now():%Y-%m-%d %H:%M})",
            f"*Hoạt động trong {hours} giờ qua*\n",
            "---",
        ]
        problems: list[str] = []

        try:
            mention_res = client.get_user_mentions(hours=hours, limit=10, context_before=2, context_after=1)
            mentions = mention_res["mentions"]
            problems.extend(mention_res["errors"])
            sections.append(f"## 🎯 1. Việc được giao & tin nhắn tag bạn ({len(mentions)})")
            if mentions:
                for m in mentions:
                    sections.append(f"### 📍 [{m['chat_name']}] — **{m['sender']}** ({m['timestamp']})")
                    for ctx in m.get("context") or []:
                        when = ctx["timestamp"][11:19]
                        marker = "👉 **" if ctx["is_mention"] else f"- *({ctx['offset']:+d}) "
                        close = " (MENTION):**" if ctx["is_mention"] else ":*"
                        sections.append(f"{marker}[{when}] {ctx['sender']}{close} {ctx['content'][:240]}")
                    sections.append("")
            else:
                sections.append("*(Không có tin nhắn nào tag bạn)*\n")
        except Mcp365Error as exc:
            sections.append(f"## 🎯 1. Việc được giao\n*(Không lấy được: {exc.message})*\n")

        try:
            feed_res = client.get_recent_feed(hours=hours, max_chats=6, limit_per_chat=4)
            problems.extend(feed_res["errors"])
            sections.append(f"## 💬 2. Thảo luận tại các nhóm ({len(feed_res['feed'])} nhóm)")
            for item in feed_res["feed"]:
                sections.append(f"### 👥 **{item['chat_name']}**")
                for msg in item["messages"][-3:]:
                    sections.append(f"- **{msg['sender']}**: {msg['content'][:150]}")
                sections.append("")
        except Mcp365Error as exc:
            sections.append(f"## 💬 2. Thảo luận\n*(Không lấy được: {exc.message})*\n")

        sections.append("## 📅 3. Lịch họp hôm nay")
        try:
            start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0).astimezone()
            events = client.get_calendar_events(start, start + timedelta(days=1))
            if events:
                for ev in events:
                    sections.append(f"- **{ev['subject']}** — {ev['start']}")
            else:
                sections.append("*(Không có cuộc họp nào)*")
        except Mcp365Error as exc:
            sections.append(f"*(Không lấy được lịch: {exc.message})*")

        sections.append("\n## 📄 4. Tài liệu SharePoint cập nhật")
        try:
            docs = sp().search_files(query="*", max_results=5)
            if docs:
                for doc in docs:
                    sections.append(
                        f"- 📄 **[{doc['title']}]({doc['path']})** *({human_size(doc['size'])}, "
                        f"{doc['modified']} bởi {doc['author']})*"
                    )
            else:
                sections.append("*(Không có tài liệu mới)*")
        except Mcp365Error as exc:
            sections.append(f"*(Không lấy được tài liệu: {exc.message})*")
            sections.append(f"> 💡 {exc.remediation}")

        return "\n".join(sections) + _errors_note(problems)


def register_resources(mcp) -> None:
    @mcp.resource("teams://chats", name="Teams conversations", mime_type="text/markdown")
    def teams_chats_resource() -> str:
        """The list of recent Teams conversations."""
        chats = teams().list_conversations(page_size=50)
        return "\n".join(f"- [{c['type']}] {c['name']} — `{c['id']}`" for c in chats)

    @mcp.resource("teams://mentions/recent", name="Recent Teams mentions", mime_type="text/markdown")
    def recent_mentions_resource() -> str:
        """Messages from the last 24 hours that mention you."""
        res = teams().get_user_mentions(hours=24, limit=15, context_before=0, context_after=0)
        if not res["mentions"]:
            return "Không có mention nào trong 24 giờ qua."
        return "\n".join(
            f"- [{m['chat_name']}] {m['sender']} ({m['timestamp']}): {m['content'][:200]}" for m in res["mentions"]
        )

    @mcp.resource("m365://health", name="Microsoft 365 connection health", mime_type="text/markdown")
    def health_resource() -> str:
        """Live status of every Microsoft 365 authentication channel."""
        return run_health_check()


def register_prompts(mcp) -> None:
    @mcp.prompt()
    def summarize_chat_thread(chat_name_or_id: str, limit: int = 60) -> str:
        """Summarise a Teams conversation: decisions, open questions and owners."""
        res = teams().get_messages(chat_name_or_id, limit=limit)
        transcript = "\n".join(f"[{m['timestamp']}] {m['sender']}: {m['content']}" for m in res["messages"])
        return (
            f"Dưới đây là {len(res['messages'])} tin nhắn từ '{res['conversation_name']}'.\n\n"
            f"Hãy tóm tắt bằng tiếng Việt theo 4 mục: (1) Quyết định đã chốt, (2) Vấn đề còn treo, "
            f"(3) Ai chịu trách nhiệm việc gì, (4) Deadline được nhắc tới. "
            f"Trích dẫn tên người và thời gian khi cần.\n\n---\n{transcript}"
        )

    @mcp.prompt()
    def draft_standup(hours: int = 24) -> str:
        """Draft a standup update from your mentions and recent activity."""
        client = teams()
        mention_res = client.get_user_mentions(hours=hours, limit=10, context_before=1, context_after=1)
        feed_res = client.get_recent_feed(hours=hours, max_chats=5, limit_per_chat=5)
        lines = ["### Tin nhắn tag bạn"]
        for m in mention_res["mentions"]:
            lines.append(f"- [{m['chat_name']}] {m['sender']}: {m['content'][:250]}")
        lines.append("\n### Thảo luận nhóm")
        for item in feed_res["feed"]:
            for msg in item["messages"][-3:]:
                lines.append(f"- [{item['chat_name']}] {msg['sender']}: {msg['content'][:200]}")
        return (
            "Dựa trên dữ liệu Teams dưới đây, hãy soạn bản standup tiếng Việt gồm 3 phần: "
            "**Hôm qua đã làm**, **Hôm nay sẽ làm**, **Vướng mắc**. "
            "Viết ngắn gọn, mỗi mục 2-4 gạch đầu dòng. Nếu thiếu thông tin thì ghi rõ là cần bổ sung, "
            "tuyệt đối không bịa.\n\n---\n" + "\n".join(lines)
        )

    @mcp.prompt()
    def triage_mentions(hours: int = 48) -> str:
        """Turn recent mentions into a prioritised action list."""
        res = teams().get_user_mentions(hours=hours, limit=20, context_before=2, context_after=2)
        blocks = []
        for m in res["mentions"]:
            ctx = "\n".join(f"    ({c['offset']:+d}) {c['sender']}: {c['content'][:200]}" for c in m.get("context") or [])
            blocks.append(
                f"- Chat: {m['chat_name']} | Người tag: {m['sender']} | {m['timestamp']}\n"
                f"  Nội dung: {m['content'][:300]}\n  Bối cảnh:\n{ctx}\n  MessageID: {m['message_id']} | ChatID: {m['chat_id']}"
            )
        return (
            "Hãy phân loại các mention Teams dưới đây thành danh sách việc cần làm, sắp xếp theo mức độ ưu tiên. "
            "Với mỗi việc, nêu: việc cần làm, người yêu cầu, deadline (nếu có), và mức ưu tiên (Cao/Trung bình/Thấp). "
            "Nếu mention chỉ mang tính thông báo thì xếp riêng vào mục 'Chỉ để biết'. "
            "Kèm ChatID và MessageID để có thể trả lời trực tiếp.\n\n---\n" + "\n\n".join(blocks)
        )


def register_all(mcp) -> None:
    register_sharepoint_tools(mcp)
    register_teams_tools(mcp)
    register_shared_tools(mcp)
    register_resources(mcp)
    register_prompts(mcp)
