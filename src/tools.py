"""Tool, resource and prompt registrations.

Declared once and mounted onto any server instance. Previously the unified
server and the two standalone servers each re-declared the same tools, and the
copies had already drifted apart (one still advertised a hardcoded user name).
"""

from __future__ import annotations

import functools
import inspect
import logging
import traceback
import urllib.parse
from datetime import datetime, timedelta
from pathlib import Path
from typing import Annotated, Any

from mcp.server.mcpserver.exceptions import ToolError
from pydantic import Field

from common import approval
from common.config import get_config
from common.errors import Mcp365Error
from common.health import run_health_check
from common.http import timeout_scope
from outlook.client import OutlookMailClient
from sharepoint import sheets
from sharepoint.client import SharePointClient, human_size
from teams.client import REACTION_EMOJI, TeamsClient, mention_label, normalize_reaction
from teams.endpoints import is_teams_media_url

logger = logging.getLogger(__name__)

# ------------------------------------------------------ per-call timeouts
#
# Every networked tool takes an optional ``timeout_seconds``. The right value
# depends on the work, so the schema text differs: a chat call that is slow
# means a sick server (and the router already fails over to another endpoint),
# while a recording download legitimately needs minutes. ``_actionable``
# applies the value to every HTTP request the tool makes; it is capped by
# ``http.timeout_max`` (600 s by default).

CHAT_TIMEOUT_DESCRIPTION = (
    "Optional per-request timeout in seconds, applied to each attempt on each Teams Chat Service endpoint "
    "(default 10 s; failed requests automatically move to the next endpoint). Normally leave it empty: a chat "
    "request that takes longer than a few seconds means the server or network is failing, not that the limit is "
    "too short. On a timeout, retry later - and for sends/edits/reactions first check whether it already went "
    "through - rather than raising this. Capped at 600."
)
TRANSFER_TIMEOUT_DESCRIPTION = (
    "Optional per-request timeout in seconds for file transfers (default 120 s; bodies are streamed, so the limit "
    "applies to each read, not the whole file). Raise it (e.g. 300-600) for very large files, meeting recordings, "
    "folders with many files, or a known-slow connection, especially after a timeout error. Capped at 600."
)
REQUEST_TIMEOUT_DESCRIPTION = (
    "Optional per-request timeout in seconds (default 30 s). Raise it only on a known-slow network after a "
    "timeout error. Capped at 600."
)

ChatTimeout = Annotated[float | None, Field(description=CHAT_TIMEOUT_DESCRIPTION)]
TransferTimeout = Annotated[float | None, Field(description=TRANSFER_TIMEOUT_DESCRIPTION)]
RequestTimeout = Annotated[float | None, Field(description=REQUEST_TIMEOUT_DESCRIPTION)]

_sp_client: SharePointClient | None = None
_teams_client: TeamsClient | None = None
_mail_client: OutlookMailClient | None = None


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


def outlook() -> OutlookMailClient:
    global _mail_client
    if _mail_client is None:
        _mail_client = OutlookMailClient()
    return _mail_client


def _actionable(fn):
    """Re-raise our typed errors as ``ToolError`` so the text survives.

    The SDK passes a ``ToolError`` message through verbatim but replaces any
    other exception with a bare "Error executing tool <name>" - which would
    discard exactly the remediation the user needs.

    It also applies the tool's ``timeout_seconds`` (when it declares one) to
    every request made during the call, so tool bodies need not pass it on.
    """
    signature = inspect.signature(fn)
    takes_timeout = "timeout_seconds" in signature.parameters

    @functools.wraps(fn)
    def wrapper(*args, **kwargs):
        seconds = None
        if takes_timeout:
            try:
                seconds = signature.bind_partial(*args, **kwargs).arguments.get("timeout_seconds")
            except TypeError:
                seconds = kwargs.get("timeout_seconds")
        try:
            with timeout_scope(seconds):
                return fn(*args, **kwargs)
        except ToolError:
            raise
        except Mcp365Error as exc:
            raise ToolError(str(exc)) from exc
        except Exception as exc:
            # An unexpected bug must still tell the agent what broke and where,
            # otherwise it only sees "Error executing tool" and retries blindly.
            logger.exception("Tool %s crashed", fn.__name__)
            raise ToolError(_describe_crash(fn.__name__, exc)) from exc

    return wrapper


_SRC_DIR = Path(__file__).resolve().parent


def _describe_crash(tool_name: str, exc: BaseException) -> str:
    """One actionable paragraph for an exception that is not one of ours.

    The location is the deepest frame in this repo's ``src/`` - the code that
    misbehaved, even when the raise came from a library - shown relative to
    ``src/`` because ``client.py`` alone names three modules.
    """
    # The first frame is _actionable's own wrapper, which caught it: skip it.
    frames = traceback.extract_tb(exc.__traceback__)[1:]
    ours = [f for f in frames if Path(f.filename).resolve().is_relative_to(_SRC_DIR)] or frames
    where = ""
    if ours:
        path = Path(ours[-1].filename).resolve()
        shown = path.relative_to(_SRC_DIR).as_posix() if path.is_relative_to(_SRC_DIR) else path.name
        where = f" tại {shown}:{ours[-1].lineno} ({ours[-1].name})"
    detail = str(exc).strip() or "(không có thông điệp)"
    return (
        f"Lỗi nội bộ trong tool `{tool_name}`: {type(exc).__name__}: {detail}{where}.\n"
        "→ Lỗi ngoài dự kiến trong MCP, không phải lỗi đăng nhập/quyền. Kiểm tra lại tham số; nếu là thao tác "
        "gửi/ghi, hãy đọc lại đích (chat, file) để xem đã thực hiện chưa trước khi thử lại. "
        "Traceback đầy đủ nằm trong log MCP."
    )


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
    def search_sharepoint_files(
        query: str, max_results: int = 20, file_extension: str = "", timeout_seconds: RequestTimeout = None
    ) -> str:
        """Search SharePoint/OneDrive documents by keyword, with optional file-type filter.

        Args:
            query: Keyword to search for (e.g. 'SYS2', 'CAN', 'Architecture').
            max_results: Maximum number of files to return.
            file_extension: Optional extension filter (e.g. 'docx', 'xlsx', 'pdf').
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
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
    def read_sharepoint_link(url: str, max_depth: int = 2, timeout_seconds: RequestTimeout = None) -> str:
        """Explore a SharePoint folder tree, or show a document's metadata and version history.

        Args:
            url: SharePoint folder or document URL.
            max_depth: How many folder levels to traverse.
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
        """
        return sp().read_link(url, max_depth=max_depth)

    @mcp.tool()
    def download_sharepoint_link(
        url_or_guid: str, target_dir: str = "", timeout_seconds: TransferTimeout = None
    ) -> str:
        """Download original binary files or a whole folder from SharePoint, with no format conversion.

        Args:
            url_or_guid: SharePoint URL, sharing link, or document UniqueId (GUID).
            target_dir: Destination directory (defaults to the configured download dir).
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        return sp().download_link(url_or_guid, target_dir=target_dir)

    @mcp.tool()
    def upload_sharepoint_file(
        local_file_path: str,
        target_folder_url_or_path: str,
        is_user_confirm: approval.UserConfirm,
        target_file_name: str = "",
        timeout_seconds: TransferTimeout = None,
    ) -> str:
        """Upload a local file to SharePoint, creating any missing parent folders.

        Ask the user before calling with is_user_confirm=true. A same-name file is replaced.
        The returned Web URL opens the file in the browser (`?web=1`: Office files in Office
        Online; .md/.txt/.pdf and others in the SharePoint viewer); the direct file link downloads.
        Only people with access to that site can open it - for anyone in the organization use
        `share_file_onedrive`.

        Args:
            local_file_path: Path to the local file.
            target_folder_url_or_path: Destination folder URL or site-relative path.
            is_user_confirm: Required. True only after the user approved this exact upload.
            target_file_name: Optional remote filename (defaults to the local name).
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
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
            f"- **Thư mục:** `{res['folder']}`\n- **Item ID:** `{res['id']}`\n"
            f"- **Link xem online:** {res['webUrl']}\n- **Link tải thẳng:** {res['fileUrl']}"
        )

    @mcp.tool()
    def share_file_onedrive(
        local_file_path: str,
        is_user_confirm: approval.UserConfirm,
        folder: str = "Shared from MCP",
        link_type: str = "view",
        target_file_name: str = "",
        timeout_seconds: TransferTimeout = None,
    ) -> str:
        """Upload a local file to YOUR OneDrive and create a link anyone in the organization can open.

        For people who cannot open the team SharePoint site (or when its storage is full): the
        file goes to your personal OneDrive (`<folder>`, created if missing; a same-name file is
        replaced), then an organization-wide sharing link is created - everyone signed in to the
        company tenant who has the link can open it; people outside cannot. The link opens in the
        browser (Office Online, or the viewer for .md/.txt/.pdf/images).

        Ask the user before calling with is_user_confirm=true: show the file, the folder and the
        link type. With false, nothing is uploaded.

        Args:
            local_file_path: Path to the local file.
            is_user_confirm: Required. True only after the user approved uploading and sharing this exact file.
            folder: Folder in your OneDrive (default "Shared from MCP").
            link_type: "view" (default, read-only) or "edit".
            target_file_name: Optional remote filename (defaults to the local name).
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files.
        """
        if link_type not in ("view", "edit"):
            raise Mcp365Error(f"link_type phải là 'view' hoặc 'edit', không phải '{link_type}'.")
        who = "xem" if link_type == "view" else "SỬA"
        approval.require_confirm(
            is_user_confirm,
            "Tải file lên OneDrive cá nhân và tạo link chia sẻ cho cả tổ chức",
            f"OneDrive của bạn › `{folder or '/'}`",
            f"`{local_file_path}` → `{target_file_name or Path(local_file_path).name}`\n"
            f"Link: mọi người trong tổ chức có link đều **{who}** được.",
        )
        res = sp().share_file_onedrive(local_file_path, folder, link_type, target_file_name)
        return (
            f"✓ Đã tải `{res['name']}` ({human_size(res['size'])}) lên OneDrive của bạn (`{res['folder']}`).\n"
            f"- **Link chia sẻ trong tổ chức ({who}):** {res['share_link']}\n"
            f"- **Link xem online (chỉ bạn và người đã có quyền):** {res['webUrl']}"
        )

    @mcp.tool()
    def read_sharepoint_sheet(
        file_url_or_guid: str,
        sheet: str = "",
        max_rows: int = 60,
        include_hidden: bool = False,
        timeout_seconds: TransferTimeout = None,
    ) -> str:
        """Read an Excel workbook on SharePoint/OneDrive: list its sheets, or show one as a table.

        Rows are numbered and columns lettered, so every value has its A1 address.

        Hidden sheets (hidden and veryHidden) are skipped by default: the author hid
        them on purpose, so they are not a source of truth. The listing still names
        them as "(ẩn, bỏ qua)", and asking for one by name returns an error saying so.

        Args:
            file_url_or_guid: File URL (any site or OneDrive), sharing link, or UniqueId.
            sheet: Sheet to show. Empty lists every sheet with its size.
            max_rows: Maximum rows to render.
            include_hidden: Also read hidden sheets. Only when the user explicitly needs one.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        drive_id, item = sp().resolve_file(file_url_or_guid)
        data = sp().read_file_bytes(drive_id, item)
        return sheets.render_sheet(
            data, sheet=sheet, max_rows=max_rows, name=item.get("name", ""), include_hidden=include_hidden
        )

    @mcp.tool()
    def compare_sharepoint_versions(
        file_a: str,
        file_b: str = "",
        version_a: str = "",
        version_b: str = "",
        timeout_seconds: TransferTimeout = None,
    ) -> str:
        """Diff two SharePoint document versions, or a local file against a SharePoint document.

        Args:
            file_a: Local path or SharePoint URL/GUID.
            file_b: Optional second file to compare against file_a.
            version_a: (When file_b is omitted) earlier version label, e.g. '1.0'.
            version_b: (When file_b is omitted) later version label, e.g. '2.0' or 'latest'.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        if file_b:
            return sp().compare_documents(file_a, file_b)
        return sp().compare_versions(file_a, version_a=version_a, version_b=version_b)

    @mcp.tool()
    def delete_sharepoint_item(
        url_or_guid: str,
        is_user_confirm: approval.UserConfirm,
        permanent: bool = False,
        timeout_seconds: RequestTimeout = None,
    ) -> str:
        """Delete a SharePoint file, or a folder with everything in it.

        Call with is_user_confirm=false first: nothing is deleted and the reply shows
        the exact path, size and item count to put to the user. By default the item goes
        to the site Recycle Bin (restorable, still counts against the site quota);
        permanent=true bypasses the bin and cannot be undone - use it only when the user
        wants the space back and a copy exists elsewhere.

        Args:
            url_or_guid: File or folder URL, or a file UniqueId.
            is_user_confirm: Required. True only after the user approved deleting this exact item.
            permanent: Bypass the Recycle Bin (default False).
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
        """
        target = sp().describe_item(url_or_guid)
        what = "thư mục" if target["is_folder"] else "file"
        count = f", {target['child_count']} mục con trực tiếp" if target["is_folder"] else ""
        mode = "**XOÁ VĨNH VIỄN** (không vào thùng rác, không khôi phục được)" if permanent else "chuyển vào thùng rác"
        approval.require_confirm(
            is_user_confirm,
            f"Xoá {what} trên SharePoint",
            f"`{target['path']}`",
            f"{what.capitalize()} `{target['name']}` ({human_size(target['size'])}{count}): {mode}",
        )
        res = sp().delete_item(target, permanent=permanent)
        done = "Đã xoá vĩnh viễn" if res["permanent"] else "Đã chuyển vào thùng rác"
        return f"✓ {done} {what} `{res['name']}` ({human_size(res['size'])}).\n- **Đường dẫn:** `{res['path']}`"

    @mcp.tool()
    def sync_folder_to_sharepoint(
        local_dir: str,
        target_folder: str,
        is_user_confirm: approval.UserConfirm,
        dry_run: bool = True,
        timeout_seconds: TransferTimeout = None,
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
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        if not dry_run:
            approval.require_confirm(
                is_user_confirm,
                "Đồng bộ thư mục lên SharePoint",
                target_folder,
                f"Tải các file mới/đổi từ `{local_dir}`",
            )
        return sp().sync_folder_up(local_dir, target_folder, dry_run=dry_run)

    @mcp.tool()
    def download_meeting_recordings(
        target_dir: str = "", limit: int = 3, query: str = "Recording", timeout_seconds: TransferTimeout = None
    ) -> str:
        """Find and download Teams meeting recordings stored in SharePoint/OneDrive.

        Args:
            target_dir: Destination directory.
            limit: Maximum number of recordings to download.
            query: Search term used to locate recordings.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        return sp().download_meeting_recordings(target_dir=target_dir, limit=limit, query=query)


def register_teams_tools(mcp) -> None:
    mcp = _ErrorAwareServer(mcp)

    @mcp.tool()
    def list_teams_chats(
        limit: int = 30, filter_keyword: str = "", chat_type: str = "", timeout_seconds: ChatTimeout = None
    ) -> str:
        """List recent Teams group chats, 1:1 chats, meeting chats and channels.

        The keyword is matched without diacritics, so ``nam son`` finds
        ``Nguyễn Phan Nam Sơn``, and it is applied across every known
        conversation before ``limit`` truncates the result.

        Args:
            limit: Maximum number of conversations to return.
            filter_keyword: Optional keyword filter on chat name or last message.
            chat_type: Optional type filter: DirectChat, GroupChat, Channel or MeetingChat.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
        """
        chats = teams().list_conversations(page_size=limit, filter_keyword=filter_keyword, chat_type=chat_type)
        if not chats:
            return "Không tìm thấy cuộc trò chuyện nào khớp."
        out = [
            f"# Cuộc trò chuyện Microsoft Teams ({len(chats)})\n",
            "| Loại | Tên | Người gửi cuối | Hoạt động | Chat ID |",
            "| --- | --- | --- | --- | --- |",
        ]
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
        timeout_seconds: ChatTimeout = None,
    ) -> str:
        """Read message history from a Teams chat, channel or 1:1 conversation.

        Args:
            chat_name_or_id: Chat name (partial match works), or thread ID.
            limit: Number of recent messages to fetch.
            since: Optional time filter: 'today', 'yesterday', '6h', '3d' or 'YYYY-MM-DD' (local time).
            only_mentions: Only return messages that mention you.
            output_file: Optional path to also save the transcript as Markdown.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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
        hours: int = 48,
        max_chats: int = 8,
        limit_per_chat: int = 8,
        filter_keyword: str = "",
        timeout_seconds: ChatTimeout = None,
    ) -> str:
        """Fetch new messages across all active chats and channels in one parallel call.

        Args:
            hours: How many hours back to look.
            max_chats: Maximum conversations to scan.
            limit_per_chat: Maximum messages per conversation.
            filter_keyword: Optional filter on chat name or content.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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
    def get_my_mentions(
        hours: int = 72,
        limit: int = 20,
        context_before: int = 2,
        context_after: int = 2,
        timeout_seconds: ChatTimeout = None,
    ) -> str:
        """Find messages that mention you, across group chats, channels and 1:1 chats.

        Matching uses the authoritative mention payload Teams attaches to each
        message (your user MRI), so it works regardless of how your display name
        is rendered. Returns surrounding messages for context.

        Args:
            hours: How many hours back to look.
            limit: Maximum mentions to return.
            context_before: Messages to include before each mention (0-10).
            context_after: Messages to include after each mention (0-10).
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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
            out.append(
                f"### 📍 [{m['chat_name']}] — **{m['sender']}** tag bạn ({m['timestamp']}) · *{m['mention_reason']}*"
            )
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
    def get_new_mentions_since(cursor: str = "", limit: int = 20, timeout_seconds: ChatTimeout = None) -> str:
        """Return only mentions newer than a cursor, for periodic polling.

        Pass the cursor returned by the previous call. An MCP stdio server cannot
        hold a background watch loop, so the caller drives the polling (for
        example from a scheduled task).

        Args:
            cursor: Timestamp from the previous call ('YYYY-MM-DD HH:MM:SS'); empty scans the last 24h.
            limit: Maximum mentions to return.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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
    def search_teams_chat_messages(keywords: list[str], limit: int = 20, timeout_seconds: ChatTimeout = None) -> str:
        """Search recent messages across chats and channels for any of several keywords.

        Args:
            keywords: Keywords or phrases to look for (e.g. ['DTC', 'S5', 'review']).
            limit: Maximum matching messages to return.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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
        mentions: list[str] | None = None,
        share_scope: str = "members",
        timeout_seconds: ChatTimeout = None,
    ) -> str:
        """Send a Teams message, optionally tagging people, quoting a message or attaching a file.

        Microsoft Teams is sensitive: ALWAYS ask the user first. Show them the exact
        message, the destination chat and who will be tagged, wait for an explicit
        yes, and only then call with is_user_confirm=true. With false, nothing is sent
        and the draft is returned for you to show them.

        RULE: in a group chat, a message meant for specific people MUST tag them via
        `mentions`. Busy groups bury untagged messages and the person never sees it.

        Args:
            chat_name_or_id: Chat name (partial match works) or thread ID.
            message: Message text; **bold**, *italic*, `code` and [links](url) are supported.
                Write `@Name` where a tag should appear; untagged-in-text people are tagged at the start.
            is_user_confirm: Required. True only after the user approved this exact message to this chat.
            reply_to_id: Optional ID (from read_teams_chat) of a message in this chat to quote-reply to. Nothing is
                sent if that message cannot be read.
            file_path: Optional local file to attach, as a link in the message that opens it in the browser.
                In a chat (1:1, group, meeting) it goes to YOUR OneDrive › "Microsoft Teams Chat Files", like
                Teams does (a same-name file is never replaced: "name (1).ext"). In a channel it goes to the
                SharePoint attachment folder (`sharepoint.attachment_folder`) as before - people with access to
                that site can open it.
            share_scope: Who can open a chat attachment: "members" (default) - only the chat's members, each
                granted by name, no invitation email; "organization" - anyone in the company with the link.
                Ignored for channels.
            mentions: People to tag: full name (diacritics optional; with "(Unit)" it must match
                exactly), email/UPN, alias ("hoangnh21") or MRI "8:orgid:<guid>". Chat history first,
                then the directory. Namesakes -> error listing candidates, nothing sent.
            timeout_seconds: Optional per-request timeout. Leave empty unless `file_path` is a large file; a slow send
                means a sick server - check whether it was sent before retrying.
        """
        if share_scope not in ("members", "organization"):
            raise Mcp365Error(f"share_scope phải là 'members' hoặc 'organization', không phải '{share_scope}'.")
        conv = teams().find_conversation(chat_name_or_id)
        people = teams().resolve_mentions(conv["id"], mentions) if mentions else []
        detail = message
        if file_path:
            if conv.get("type") == "Channel" or "@thread.tacv2" in conv["id"]:
                where = f"SharePoint › `{get_config().sharepoint.attachment_folder}`"
                who = "người có quyền trên site SharePoint đó"
            else:
                where = "OneDrive của bạn › `Microsoft Teams Chat Files`"
                who = ("chỉ các thành viên của chat này (cấp quyền từng người, không gửi email mời)"
                       if share_scope == "members" else "mọi người trong tổ chức có link")
            detail += f"\n\n**Đính kèm:** `{file_path}` → {where}\n**Ai mở được file:** {who}"
        if people:
            detail += "\n\n**Tag:** " + ", ".join(mention_label(p) for p in people)
        approval.require_confirm(is_user_confirm, "Gửi tin nhắn Teams", f"{conv['name']} (`{conv['id']}`)", detail)
        return _render_send(
            teams().send_message(
                conversation_id_or_name=conv["id"],
                message=message,
                reply_to_id=reply_to_id or None,
                file_path=file_path or None,
                mentions=people,
                share_scope=share_scope,
            )
        )

    def _render_send(res: dict[str, Any]) -> str:
        extra = []
        if res.get("reply_to_id"):
            extra.append(f"- **Trả lời tin nhắn:** `{res['reply_to_id']}`")
        if res.get("attached_file"):
            f = res["attached_file"]
            extra.append(f"- **File đính kèm:** [{f['name']}]({f.get('share_link') or f['webUrl']})"
                         + (f" · lưu ở {f['location']}" if f.get("location") else ""))
        if res.get("mentioned"):
            extra.append("- **Đã tag:** " + ", ".join(res["mentioned"]))
        if res.get("message_id"):
            extra.append(f"- **Message ID:** `{res['message_id']}`")
        suffix = "\n" + "\n".join(extra) if extra else ""
        return f"✓ Đã gửi tin nhắn tới **{res['conversation_name']}** (`{res['conversation_id']}`):{suffix}\n\n> {res['message_sent']}"

    @mcp.tool()
    def reply_to_channel_thread(
        channel_name_or_id: str,
        parent_message_id: str,
        message: str,
        is_user_confirm: approval.UserConfirm,
        timeout_seconds: ChatTimeout = None,
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
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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
        chat_name_or_id: str,
        message_id: str,
        new_message: str,
        is_user_confirm: approval.UserConfirm,
        mentions: list[str] | None = None,
        timeout_seconds: ChatTimeout = None,
    ) -> str:
        """Edit one of your own previously sent Teams messages, keeping or setting its tags.

        ALWAYS ask the user first. Show them the new text, the chat and who will be tagged,
        and edit only after an explicit yes. With is_user_confirm=false nothing is changed
        and the draft is returned for you to show them.

        Tags: an edit replaces the whole message, so `@Name` written as plain text is NOT a tag.
        - `mentions` given: tag exactly those people, same rules as `send_teams_message`
          (name, "Name (Unit)", email/UPN, alias or MRI; namesakes are refused, not guessed).
          Nothing is edited if a name is not found or is ambiguous.
        - `mentions` omitted: the original message is read and each person it tagged stays
          tagged if the new text still writes `@` + their name (full display name, or without
          the "(Org unit)" suffix). Tags whose `@Name` was removed are dropped; the draft
          lists both.
        - `mentions=[]`: tag nobody.

        Args:
            chat_name_or_id: Chat name or thread ID.
            message_id: ID of the message to edit.
            new_message: Replacement text. Write `@Name` where a tag should appear; people in
                `mentions` not written in the text are tagged at the start.
            is_user_confirm: Required. True only after the user approved this exact new text and tags.
            mentions: Optional people to tag (name, email, alias or MRI), e.g. ["Nguyễn Minh Dân"]. Omit to keep the
                original's tags (see above); [] removes all tags.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
        """
        conv = teams().find_conversation(chat_name_or_id)
        dropped: list[str] = []
        if mentions is None:
            people, dropped = teams().mentions_to_keep(conv["id"], message_id, new_message)
            label = "Tag (giữ từ tin gốc)"
        else:
            people = teams().resolve_mentions(conv["id"], mentions) if mentions else []
            label = "Tag"
        detail = new_message
        if people:
            detail += f"\n\n**{label}:** " + ", ".join(mention_label(p) for p in people)
        if dropped:
            detail += "\n\n**Bỏ tag (không còn `@Tên` trong nội dung mới):** " + ", ".join(dropped)
        approval.require_confirm(is_user_confirm, "Sửa tin nhắn Teams", f"{conv['name']} · tin `{message_id}`", detail)
        res = teams().edit_message(conv["id"], message_id=message_id, new_message=new_message, mentions=people)
        tagged = f"\n- **Đã tag:** {', '.join(res['mentioned'])}" if res.get("mentioned") else ""
        return f"✓ Đã sửa tin nhắn `{res['message_id']}` trong '{res['conversation_name']}':{tagged}\n{res['new_message']}"

    @mcp.tool()
    def delete_teams_message(
        chat_name_or_id: str,
        message_id: str,
        is_user_confirm: approval.UserConfirm,
        timeout_seconds: ChatTimeout = None,
    ) -> str:
        """Delete (recall) one of your own previously sent Teams messages.

        ALWAYS ask the user first and delete only after an explicit yes.

        Args:
            chat_name_or_id: Chat name or thread ID.
            message_id: ID of the message to delete.
            is_user_confirm: Required. True only after the user approved deleting this message.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
        """
        conv = teams().find_conversation(chat_name_or_id)
        approval.require_confirm(
            is_user_confirm, "Xoá tin nhắn Teams", f"{conv['name']} (`{conv['id']}`)", f"Xoá tin nhắn `{message_id}`"
        )
        res = teams().delete_message(conv["id"], message_id=message_id)
        return f"✓ Đã xoá tin nhắn `{res['message_id']}` khỏi '{res['conversation_name']}'."

    @mcp.tool()
    def react_to_teams_message(
        chat_name_or_id: str,
        message_id: str,
        reaction: str,
        is_user_confirm: approval.UserConfirm,
        remove: bool = False,
        timeout_seconds: ChatTimeout = None,
    ) -> str:
        """Add or remove a reaction on a Teams message instead of replying.

        Prefer this when the message only needs acknowledgement—for example,
        someone confirms that requested work is complete and no follow-up
        question remains. Reactions are visible communication, so ALWAYS show
        the exact reaction, chat and message ID, then wait for explicit approval.

        Args:
            chat_name_or_id: Chat name, channel name or thread ID.
            message_id: ID of the message to react to.
            reaction: like, heart, laugh, surprised, sad, angry, or the matching emoji.
            is_user_confirm: Required. True only after approval of this exact reaction and message.
            remove: True to remove your matching reaction instead of adding it.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
        """
        reaction_key = normalize_reaction(reaction)
        emoji = REACTION_EMOJI[reaction_key]
        conv = teams().find_conversation(chat_name_or_id)
        action = "Gỡ reaction Teams" if remove else "Thả reaction Teams"
        preview = f"{'Gỡ' if remove else 'Thả'} {emoji} `{reaction_key}` trên tin nhắn `{message_id}`"
        approval.require_confirm(
            is_user_confirm,
            action,
            f"{conv['name']} (`{conv['id']}`)",
            preview,
        )
        res = teams().react_to_message(
            conv["id"],
            message_id=message_id,
            reaction=reaction_key,
            remove=remove,
        )
        verb = "Đã gỡ" if remove else "Đã thả"
        return (
            f"✓ {verb} {res['emoji']} `{res['reaction']}` trên tin nhắn `{res['message_id']}` "
            f"trong **{res['conversation_name']}**."
        )

    @mcp.tool()
    def download_chat_attachments(
        chat_name_or_id: str,
        target_dir: str = "",
        limit: int = 5,
        file_name: str = "",
        scan_messages: int = 50,
        timeout_seconds: TransferTimeout = None,
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
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        res = teams().get_messages(chat_name_or_id, limit=scan_messages)
        wanted = file_name.lower().strip()
        links: list[str] = []
        for msg in reversed(res["messages"]):
            candidates = [(a["name"], a["url"]) for a in msg.get("attachments", [])]
            candidates += [(link.rsplit("/", 1)[-1], link) for link in msg.get("sharepoint_links", [])]
            candidates += [(img["name"], img["url"]) for img in msg.get("images", [])]
            for name, link in candidates:
                if wanted and wanted not in urllib.parse.unquote(name).lower():
                    continue
                if (name, link) not in links:
                    links.append((name, link))
        if not links:
            what = f"file khớp '{file_name}'" if wanted else "file đính kèm hay link SharePoint/OneDrive nào"
            return f"Không tìm thấy {what} trong {scan_messages} tin gần nhất của '{res['conversation_name']}'."

        reports, failures = [], []
        for name, link in links[:limit]:
            try:
                if is_teams_media_url(link):
                    dest = Path(target_dir or "downloads").expanduser().resolve()
                    dest.mkdir(parents=True, exist_ok=True)
                    out_p = teams().download_image(link, dest / name)
                    reports.append(f"✓ Đã tải ảnh: `{out_p}` ({human_size(out_p.stat().st_size)})")
                else:
                    reports.append(sp().download_link(link, target_dir=target_dir))
            except Mcp365Error as exc:
                failures.append(f"- `{link[:70]}…`: {exc.message}")
        body = f"# Đã xử lý {len(reports)}/{min(len(links), limit)} tệp từ '{res['conversation_name']}'\n\n"
        body += "\n\n---\n\n".join(reports)
        if failures:
            body += "\n\n> ⚠️ **Thất bại:**\n" + "\n".join(f"> {f}" for f in failures)
        return body

    @mcp.tool()
    def download_message_images(
        chat_name_or_id: str,
        message_id: str = "",
        target_dir: str = "",
        limit: int = 5,
        timeout_seconds: TransferTimeout = None,
    ) -> str:
        """Download inline screenshots and image attachments from Teams chat messages.

        Downloads images to a local directory so agents and tools can inspect them.
        If `message_id` is specified, downloads all images from that exact message.
        Otherwise, downloads recent images from the conversation.

        Args:
            chat_name_or_id: Chat name or thread ID.
            message_id: Optional exact message ID to download images from.
            target_dir: Local directory to save images (defaults to downloads/images).
            limit: Maximum number of images to download (default 5, max 20).
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        downloaded = teams().download_message_images(
            chat_name_or_id, message_id=message_id, target_dir=target_dir, limit=limit
        )
        if not downloaded:
            where = f"trong tin nhắn `{message_id}`" if message_id else "gần đây"
            return f"Không tìm thấy hình ảnh nào {where} trong cuộc trò chuyện."

        lines = [
            f"# 🖼️ Đã tải {len(downloaded)} hình ảnh thành công:\n",
        ]
        for img in downloaded:
            lines.append(
                f"- **{img['name']}** ({human_size(img['size'])})\n"
                f"  - Đường dẫn local: `{img['path']}`\n"
                f"  - Từ message ID: `{img['message_id']}` ({img['sender']} · {img['timestamp']})"
            )
        return "\n".join(lines)

    @mcp.tool()
    def find_user(query: str, max_results: int = 5, timeout_seconds: RequestTimeout = None) -> str:
        """Find a colleague in Microsoft 365 / Teams by name, email, alias, phone or keyword.

        Searches the organization's directory and returns contact info (email, phone,
        job title, department), Teams MRI (for @mentioning), and direct 1:1 chat ID.

        Args:
            query: Name (with or without diacritics, e.g. "Trịnh Anh Tuấn", "nam son"), email, alias ("tuanta81"), phone, or keyword.
            max_results: Maximum number of people to return (default 5, max 20).
            timeout_seconds: Total time for the whole search (default 30 s). A slow or failing source is skipped
                and the next one tried; when time runs out the reply says so and shows what was found.
        """
        results, notes = teams().search_users_bounded(query, max_results=max_results, budget=timeout_seconds)
        skipped = ("\n\n> ⚠️ Nguồn bị bỏ qua: " + "; ".join(notes)) if notes else ""
        if not results:
            return f"Không tìm thấy người nào khớp với từ khóa: '{query}'.{skipped}"

        lines = [
            f"# 👤 Kết quả tìm kiếm người: `{query}` ({len(results)} người)\n",
        ]
        for idx, p in enumerate(results, 1):
            lines.append(f"### {idx}. {p['name']}")
            if p.get("job_title") or p.get("department"):
                lines.append(f"- **Chức vụ / Phòng ban:** {p.get('job_title') or 'N/A'} · {p.get('department') or 'N/A'}")
            if p.get("email") or p.get("upn"):
                lines.append(f"- **Email:** `{p.get('email') or p.get('upn')}`" + (f" (UPN: `{p['upn']}`)" if p.get('upn') and p['upn'] != p.get('email') else ""))
            if p.get("phone"):
                lines.append(f"- **Điện thoại:** `{p['phone']}`")
            if p.get("office"):
                lines.append(f"- **Văn phòng:** {p['office']}")
            if p.get("teams_mri"):
                lines.append(f"- **Teams MRI (để tag):** `{p['teams_mri']}`")
            if p.get("direct_chat_id"):
                lines.append(f"- **Chat 1:1 ID:** `{p['direct_chat_id']}`")
            lines.append("")

        lines.append("> Mẹo: Dùng tên này trong `send_teams_message(mentions=[...])` để tag, hoặc dùng Chat 1:1 ID để gửi tin nhắn riêng.")
        return "\n".join(lines).strip() + skipped

    @mcp.tool()
    def get_calendar_today(days: int = 1, timeout_seconds: RequestTimeout = None) -> str:
        """List your Teams calendar meetings, with join links.

        Uses the Teams middle-tier session from Chrome, because the Azure CLI
        Graph token carries no Calendars.* scope.

        Args:
            days: How many days ahead to include (1 = today only).
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
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


def register_mail_tools(mcp) -> None:
    mcp = _ErrorAwareServer(mcp)

    @mcp.tool()
    def list_emails(
        folder: str = "inbox",
        limit: int = 20,
        unread_only: bool = False,
        since: str = "",
        query: str = "",
        timeout_seconds: RequestTimeout = None,
    ) -> str:
        """List recent or searched Outlook email from one mailbox folder.

        Args:
            folder: inbox, sent, drafts, deleted, archive, junk, or an Outlook folder ID.
            limit: Maximum messages to return (1-50).
            unread_only: Return only unread messages.
            since: Optional YYYY-MM-DD or ISO 8601 received-time lower bound.
            query: Optional Outlook mail search text.
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
        """
        messages = outlook().list_messages(
            folder=folder, limit=limit, unread_only=unread_only, since=since, query=query
        )
        if not messages:
            return f"Không có email nào khớp trong thư mục `{folder}`."
        out = [f"# Outlook mail · {folder} ({len(messages)} email)\n"]
        for msg in messages:
            state = "📩 chưa đọc" if not msg["is_read"] else "đã đọc"
            attachment = " · 📎 có file" if msg["has_attachments"] else ""
            out.append(f"## {msg['subject']}")
            out.append(f"- **Từ:** {msg['from']} · **Nhận:** {msg['received']} · {state}{attachment}")
            out.append(f"- **Message ID:** `{msg['id']}`")
            if msg["preview"]:
                out.append(f"> {msg['preview'][:500]}")
            out.append("")
        out.append("> Gọi `read_email(message_id)` để đọc toàn bộ nội dung.")
        return "\n".join(out)

    @mcp.tool()
    def read_email(message_id: str, timeout_seconds: RequestTimeout = None) -> str:
        """Read one Outlook email in full using an ID returned by `list_emails`.

        Args:
            message_id: Outlook message ID.
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
        """
        msg = outlook().get_message(message_id)
        out = [
            f"# {msg['subject']}",
            f"- **Từ:** {msg['from']}",
            f"- **Tới:** {', '.join(msg['to']) or '(trống)'}",
        ]
        if msg["cc"]:
            out.append(f"- **CC:** {', '.join(msg['cc'])}")
        out.extend(
            [
                f"- **Nhận:** {msg['received']}",
                f"- **Message ID:** `{msg['id']}`",
                f"- **Đính kèm:** {'Có' if msg['has_attachments'] else 'Không'}",
                "",
                "---",
                "",
                msg.get("body") or msg["preview"] or "(email không có nội dung)",
            ]
        )
        if msg["web_link"]:
            out.append(f"\n---\n[Mở trong Outlook]({msg['web_link']})")
        return "\n".join(out)

    @mcp.tool()
    def send_email(
        to: list[str],
        subject: str,
        body: str,
        is_user_confirm: approval.UserConfirm,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        timeout_seconds: RequestTimeout = None,
    ) -> str:
        """Send a plain-text Outlook email after per-message human approval.

        ALWAYS show the user the exact To/CC/BCC, subject and body first. Call
        with false to receive that draft back; call with true only after the
        user explicitly approves this exact email. Approval for another message
        or a general instruction to send does not count.

        Args:
            to: Recipient email addresses.
            subject: Exact email subject.
            body: Exact plain-text email body.
            is_user_confirm: Required. True only after the user approved this exact email and recipients.
            cc: Optional CC recipient addresses.
            bcc: Optional BCC recipient addresses.
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
        """
        to = [address.strip() for address in to if address.strip()]
        cc = [address.strip() for address in cc or [] if address.strip()]
        bcc = [address.strip() for address in bcc or [] if address.strip()]
        target = f"To: {', '.join(to) or '(trống)'}"
        if cc:
            target += f"\nCC: {', '.join(cc)}"
        if bcc:
            target += f"\nBCC: {', '.join(bcc)}"
        preview = f"**Subject:** {subject}\n\n{body}"
        approval.require_confirm(is_user_confirm, "Gửi email Outlook", target, preview)
        result = outlook().send_message(to=to, subject=subject, body=body, cc=cc, bcc=bcc)
        copied = f"\n- **CC:** {', '.join(result['cc'])}" if result["cc"] else ""
        blind = f"\n- **BCC:** {', '.join(result['bcc'])}" if result["bcc"] else ""
        return (
            "✓ Outlook Web đã nhận email để gửi (HTTP 202; chưa phải xác nhận phát thành công).\n"
            f"- **To:** {', '.join(result['to'])}{copied}{blind}\n"
            f"- **Subject:** {result['subject']}\n- **Lưu:** Sent Items"
        )


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
    def extract_action_items(hours: int = 72, limit: int = 25, timeout_seconds: ChatTimeout = None) -> str:
        """Collect messages that look like assigned work, as structured raw material.

        Returns mentions plus request-shaped messages with their chat, sender,
        timestamp and message ID, so they can be triaged into a task list. This
        tool does no summarising of its own - it gathers the evidence.

        Args:
            hours: How many hours back to scan.
            limit: Maximum items to return.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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
            out.append(
                f"- **Nguồn:** `{item['source']}` · **Message ID:** `{item['message_id']}` · **Chat ID:** `{item['chat_id']}`"
            )
            out.append(f"> {item['content'][:500]}")
            for ctx in item.get("context") or []:
                if not ctx["is_mention"]:
                    out.append(f"  - *({ctx['offset']:+d}) {ctx['sender']}:* {ctx['content'][:160]}")
            out.append("")
        errs = mention_res["errors"] + keyword_res["errors"]
        return "\n".join(out) + _errors_note(errs)

    @mcp.tool()
    def get_daily_briefing(hours: int = 24, timeout_seconds: ChatTimeout = None) -> str:
        """Morning briefing: mentions, 1:1 messages, active discussions, calendar and recent documents.

        Args:
            hours: How many hours back to synthesise.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
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

        def feed_lines(feed: list[dict[str, Any]], icon: str) -> list[str]:
            lines = []
            for item in feed:
                lines.append(f"### {icon} **{item['chat_name']}**")
                lines.extend(f"- **{msg['sender']}**: {msg['content'][:150]}" for msg in item["messages"][-3:])
                lines.append("")
            return lines

        # A 1:1 message that tags nobody reaches neither the mention scan nor
        # the group feed; this section is the only place it shows up.
        try:
            dm_res = client.get_recent_feed(
                hours=hours, max_chats=8, limit_per_chat=4, chat_types=("DirectChat",), incoming_only=True
            )
            problems.extend(dm_res["errors"])
            sections.append(f"## 📨 2. Tin nhắn 1:1 ({len(dm_res['feed'])} cuộc trò chuyện)")
            sections.extend(feed_lines(dm_res["feed"], "👤") or ["*(Không có tin nhắn 1:1 mới)*\n"])
        except Mcp365Error as exc:
            sections.append(f"## 📨 2. Tin nhắn 1:1\n*(Không lấy được: {exc.message})*\n")

        try:
            feed_res = client.get_recent_feed(hours=hours, max_chats=6, limit_per_chat=4)
            problems.extend(feed_res["errors"])
            sections.append(f"## 💬 3. Thảo luận tại các nhóm ({len(feed_res['feed'])} nhóm)")
            sections.extend(feed_lines(feed_res["feed"], "👥"))
        except Mcp365Error as exc:
            sections.append(f"## 💬 3. Thảo luận\n*(Không lấy được: {exc.message})*\n")

        sections.append("## 📅 4. Lịch họp hôm nay")
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

        sections.append("\n## 📄 5. Tài liệu SharePoint cập nhật")
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
            ctx = "\n".join(
                f"    ({c['offset']:+d}) {c['sender']}: {c['content'][:200]}" for c in m.get("context") or []
            )
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
    register_mail_tools(mcp)
    register_shared_tools(mcp)
    register_resources(mcp)
    register_prompts(mcp)
