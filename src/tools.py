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
from common.multi import as_list, local_files, run_ordered, same_names
from outlook import client as mail_mod
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


def _file_kwargs(file_path: str, file_paths: Any, attached: list[Path]) -> dict[str, Any]:
    """Client kwargs for the files: the old ``file_path=`` call when there is one file, else ``file_paths``."""
    if not attached:
        return {"file_path": None}
    if len(attached) == 1 and not as_list(file_paths):
        return {"file_path": file_path}
    return {"file_paths": [str(p) for p in attached]}


def _batch_files(one: str, many: Any, target_file_name: str) -> list[Path]:
    """Files for a multi-file upload: all present, no two with the same name, no rename."""
    if target_file_name:
        raise Mcp365Error(
            "`target_file_name` chỉ dùng khi tải một file.", "Bỏ `target_file_name`, hoặc tải file đó riêng một lần."
        )
    files = local_files(one, many, what="file cần tải lên")
    if dupes := same_names(files):
        raise Mcp365Error(
            f"Tên file trùng nhau ({', '.join(dupes)}): cùng một thư mục đích, file sau sẽ ghi đè file trước.",
            "Đổi tên cho khác nhau, hoặc tải từng file vào thư mục khác nhau.",
        )
    return files


def _batch_draft(files: list[Path]) -> str:
    return "\n".join(f"{i}. `{p}` → `{p.name}` · {human_size(p.stat().st_size)}" for i, p in enumerate(files, 1))


def _batch_upload(files: list[Path], one, render, where: str) -> str:
    """Upload ``files`` in parallel; report each in input order, failures listed, never hidden.

    Unlike a message, these uploads are independent: the ones that worked stay and are reported.
    Nothing worked -> error.
    """
    outcomes = run_ordered(one, files, key=lambda p: p.name.casefold())
    for _res, exc in outcomes:
        if exc is not None and not isinstance(exc, Mcp365Error):
            raise exc
    done = [res for res, exc in outcomes if exc is None]
    failed = [f"`{p}`: {exc.message}" for p, (_res, exc) in zip(files, outcomes, strict=True) if exc is not None]
    if not done:
        remedies = dict.fromkeys(exc.remediation for _res, exc in outcomes if exc is not None and exc.remediation)
        raise Mcp365Error(
            f"Không tải được file nào {where} ({len(files)} file):\n" + "\n".join(f"- {f}" for f in failed),
            " | ".join(remedies),
        )
    out = [f"# Đã tải {len(done)}/{len(files)} file {where}\n"] + [render(res) for res in done]
    if failed:
        out.append(f"\n> ⚠️ **Thất bại ({len(failed)}/{len(files)}):**\n" + "\n".join(f"> - {f}" for f in failed))
    return "\n".join(out)


def _messages_by_id(client, chat: str, ids: list[str], scan: int, include_quoted: bool):
    """``(chat name, messages, errors)`` for several message IDs: one history read, then by ID for the rest.

    ``messages`` is ordered so that ``reversed(messages)`` walks the IDs in the order given.
    """
    extra = {"resolve_quotes": True} if include_quoted else {}
    res = client.get_messages(chat, limit=scan, **extra)
    by_id = {str(m.get("id")): m for m in res["messages"]}
    found, errors = [], []
    for mid in ids:
        if mid in by_id:
            found.append(by_id[mid])
            continue
        try:
            found.append(client.get_message(res["conversation_id"], mid, scan=1, resolve_quotes=include_quoted)[1])
        except Mcp365Error as exc:
            errors.append(f"`{mid}`: {exc.message}")
    if not found:
        raise Mcp365Error(
            f"Không đọc được tin nào trong '{res['conversation_name']}':\n" + "\n".join(f"- {e}" for e in errors),
            "Kiểm tra lại message ID (lấy từ `read_teams_chat`) và chat chứa nó.",
        )
    return res["conversation_name"], list(reversed(found)), errors


def _errors_note(errors: list[str]) -> str:
    """Render partial failures instead of silently returning a short answer."""
    if not errors:
        return ""
    lines = "\n".join(f"> - {e}" for e in errors[:10])
    more = f"\n> - …và {len(errors) - 10} lỗi khác" if len(errors) > 10 else ""
    return f"\n\n> ⚠️ **Không quét được {len(errors)} cuộc trò chuyện** (kết quả bên trên chưa đầy đủ):\n{lines}{more}"


def _quoted_files_note(msg: dict[str, Any]) -> str:
    """Tên + link file của tin được trích/chuyển tiếp (dòng trích dẫn trong nội dung chỉ có tên).

    File trùng với file của chính tin (tin chuyển tiếp mang theo file) thì không nhắc lại.
    """
    own = {f["url"] for f in msg.get("attachments") or []}
    lines = []
    for item in msg.get("quoted") or []:
        files = [f for f in item.get("attachments") or [] if f["url"] not in own]
        if files:
            which = f"tin {item['message_id']}" if item.get("message_id") else "tin gốc"
            links = ", ".join(f"[{f['name']}]({f['url']})" for f in files)
            lines.append(f"\n  ↩️📎 File của {which}: {links}")
    return "".join(lines)


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
        suffix += _quoted_files_note(msg)
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
        target_folder_url_or_path: str,
        is_user_confirm: approval.UserConfirm,
        local_file_path: str = "",
        target_file_name: str = "",
        timeout_seconds: TransferTimeout = None,
        local_file_paths: list[str] | str | None = None,
    ) -> str:
        """Upload one or more local files to a SharePoint folder, creating any missing parent folders.

        Ask the user before calling with is_user_confirm=true. A same-name file is replaced.
        The returned Web URL opens the file in the browser (`?web=1`: Office files in Office
        Online; .md/.txt/.pdf and others in the SharePoint viewer); the direct file link downloads.
        Only people with access to that site can open it - for anyone in the organization use
        `share_file_onedrive`.

        Args:
            target_folder_url_or_path: Destination folder URL or site-relative path.
            is_user_confirm: Required. True only after the user approved this exact upload.
            local_file_path: Path to the local file.
            target_file_name: Optional remote filename (defaults to the local name). Only with a single file.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
            local_file_paths: Several local files for the same folder, in one approval (combined with
                `local_file_path`, that one first). Uploaded in parallel; each file's result is reported, and a
                failed file does not stop the others. Two files with the same name are refused (one would
                replace the other). For a whole folder use `sync_folder_to_sharepoint`.
        """
        if as_list(local_file_paths):
            files = _batch_files(local_file_path, local_file_paths, target_file_name)
            approval.require_confirm(
                is_user_confirm,
                f"Tải {len(files)} file lên SharePoint",
                target_folder_url_or_path,
                _batch_draft(files),
            )
            client = sp()  # built once, before the worker threads
            return _batch_upload(
                files,
                lambda p: client.upload_file(str(p), target_folder_url_or_path, None),
                lambda res: (f"✓ `{res['name']}` ({human_size(res['size'])}) · Item ID `{res['id']}`\n"
                             f"  - Link xem online: {res['webUrl']}\n  - Link tải thẳng: {res['fileUrl']}"),
                f"lên SharePoint (`{target_folder_url_or_path}`)",
            )
        if not local_file_path:
            raise Mcp365Error("Thiếu file cần tải lên.", "Truyền `local_file_path` (một file) hoặc `local_file_paths`.")
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
        is_user_confirm: approval.UserConfirm,
        local_file_path: str = "",
        folder: str = "Shared from MCP",
        link_type: str = "view",
        target_file_name: str = "",
        timeout_seconds: TransferTimeout = None,
        local_file_paths: list[str] | str | None = None,
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
            is_user_confirm: Required. True only after the user approved uploading and sharing this exact file.
            local_file_path: Path to the local file.
            folder: Folder in your OneDrive (default "Shared from MCP").
            link_type: "view" (default, read-only) or "edit".
            target_file_name: Optional remote filename (defaults to the local name). Only with a single file.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files.
            local_file_paths: Several local files, one link each, in one approval (combined with
                `local_file_path`, that one first). Uploaded in parallel; each file's result is reported, and a
                failed file does not stop the others. Two files with the same name are refused.
        """
        if link_type not in ("view", "edit"):
            raise Mcp365Error(f"link_type phải là 'view' hoặc 'edit', không phải '{link_type}'.")
        who = "xem" if link_type == "view" else "SỬA"
        if as_list(local_file_paths):
            files = _batch_files(local_file_path, local_file_paths, target_file_name)
            approval.require_confirm(
                is_user_confirm,
                f"Tải {len(files)} file lên OneDrive cá nhân và tạo link chia sẻ cho cả tổ chức",
                f"OneDrive của bạn › `{folder or '/'}`",
                _batch_draft(files) + f"\nLink (mỗi file một link): mọi người trong tổ chức có link đều **{who}** được.",
            )
            client = sp()  # built once, before the worker threads
            return _batch_upload(
                files,
                lambda p: client.share_file_onedrive(str(p), folder, link_type, ""),
                lambda res: (f"✓ `{res['name']}` ({human_size(res['size'])})\n"
                             f"  - Link chia sẻ trong tổ chức ({who}): {res['share_link']}\n"
                             f"  - Link xem online (chỉ bạn và người đã có quyền): {res['webUrl']}"),
                f"lên OneDrive của bạn (`{folder or '/'}`)",
            )
        if not local_file_path:
            raise Mcp365Error("Thiếu file cần chia sẻ.", "Truyền `local_file_path` (một file) hoặc `local_file_paths`.")
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
        chat_name_or_id: str = "",
        limit: int = 30,
        since: str = "",
        only_mentions: bool = False,
        output_file: str = "",
        timeout_seconds: ChatTimeout = None,
        chat_names_or_ids: list[str] | str | None = None,
    ) -> str:
        """Read message history from one or more Teams chats, channels or 1:1 conversations.

        A quote reply or forward shows the original on its own line: message ID, sender, time, the chat it lives
        in (when not this one), its text and its files (names here, links below the message). Originals missing
        from the page are read by ID (at most a few per call); if that fails the old preview stays.

        Args:
            chat_name_or_id: Chat name (partial match works), or thread ID.
            limit: Number of recent messages to fetch.
            since: Optional time filter: 'today', 'yesterday', '6h', '3d' or 'YYYY-MM-DD' (local time).
            only_mentions: Only return messages that mention you.
            output_file: Optional path to also save the transcript as Markdown.
            timeout_seconds: Optional per-request timeout. Leave empty: a slow chat request means a sick server
                (requests already fail over between endpoints), not a short limit.
            chat_names_or_ids: Several chats read in parallel in one call (combined with `chat_name_or_id`, that
                one first); one section per chat in the order given, the same `limit`/`since`/`only_mentions`
                for each. A chat that cannot be read is listed at the end; the others are still returned.
        """
        chats = as_list(chat_name_or_id, chat_names_or_ids)
        if len(chats) > 1:
            return _read_many_chats(chats, limit, since, only_mentions, output_file)
        if not chats:
            raise Mcp365Error("Thiếu chat cần đọc.", "Truyền `chat_name_or_id` (một chat) hoặc `chat_names_or_ids`.")
        res = teams().get_messages(
            chat_name_or_id or chats[0], limit=limit, since=since or None, only_mentions=only_mentions,
            resolve_quotes=True,
        )
        if not res["messages"]:
            return f"Không có tin nhắn nào trong '{res['conversation_name']}' khớp điều kiện."
        rendered = _render_chat(res)
        if output_file:
            path = Path(output_file).expanduser().resolve()
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(rendered, encoding="utf-8")
            return f"✓ Đã lưu {len(res['messages'])} tin nhắn vào `{output_file}`.\n\n{rendered}"
        return rendered

    def _read_many_chats(chats: list[str], limit: int, since: str, only_mentions: bool, output_file: str) -> str:
        client = teams()  # built once, before the worker threads

        def read(chat: str) -> dict[str, Any]:
            return client.get_messages(
                chat, limit=limit, since=since or None, only_mentions=only_mentions, resolve_quotes=True
            )

        outcomes = run_ordered(read, chats, max_workers=get_config().http.max_workers)
        sections, errors, count = [], [], 0
        for chat, (res, exc) in zip(chats, outcomes, strict=True):
            if exc is not None:
                if not isinstance(exc, Mcp365Error):
                    raise exc
                errors.append(f"{chat}: {exc.message}")
            elif res["messages"]:
                count += len(res["messages"])
                sections.append(_render_chat(res))
            else:
                sections.append(f"# Chat: {res['conversation_name']}\n\nKhông có tin nhắn nào khớp điều kiện.")
        if not sections:
            raise Mcp365Error(
                f"Không đọc được chat nào ({len(chats)}):\n" + "\n".join(f"- {e}" for e in errors),
                "Kiểm tra lại tên/ID chat (xem `list_teams_chats`).",
            )
        rendered = "\n\n---\n\n".join(sections)
        if errors:
            rendered += (f"\n\n> ⚠️ **Không đọc được {len(errors)}/{len(chats)} chat** (kết quả trên chưa đầy đủ):\n"
                         + "\n".join(f"> - {e}" for e in errors))
        if output_file:
            path = Path(output_file).expanduser().resolve()
            path.parent.mkdir(parents=True, exist_ok=True)
            path.write_text(rendered, encoding="utf-8")
            return f"✓ Đã lưu {count} tin nhắn của {len(sections)} chat vào `{output_file}`.\n\n{rendered}"
        return rendered

    def _render_chat(res: dict[str, Any]) -> str:
        out = [
            f"# Chat: {res['conversation_name']}",
            f"- **Loại:** {res['conversation_type']}",
            f"- **Thread ID:** `{res['conversation_id']}`",
            f"- **Số tin nhắn:** {len(res['messages'])}\n",
            "---",
        ]
        out.extend(_render_messages(res["messages"], bullet=False))
        return "\n".join(out)

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
        file_paths: list[str] | str | None = None,
        allow_org_fallback: bool = False,
    ) -> str:
        """Send a Teams message, optionally tagging people, quoting a message or attaching one or more files.

        Microsoft Teams is sensitive: ALWAYS ask the user first. Show them the exact
        message, the destination chat and who will be tagged, wait for an explicit
        yes, and only then call with is_user_confirm=true. With false, nothing is sent
        and the draft is returned for you to show them.

        RULE: in a group chat, a message meant for specific people MUST tag them via
        `mentions`. Busy groups bury untagged messages and the person never sees it.

        Args:
            chat_name_or_id: Chat name (partial match works) or thread ID.
            message: Message text in markdown: **bold**, *italic*, `code`, [links](url), "- " bullets, "1. " lists
                and tables (as in mail `body_format="markdown"`); raw HTML is shown as text. Bare http(s)://
                URLs become clickable links (not inside `code`); a trailing "." "," ")" stays outside the link.
                Write `@Name` where a tag should appear; untagged-in-text people are tagged at the start.
            is_user_confirm: Required. True only after the user approved this exact message to this chat.
            reply_to_id: Optional ID (from read_teams_chat) of a message in this chat to quote-reply to. If it
                cannot be read now, the copy from an earlier read_teams_chat/mention read is quoted; with neither,
                nothing is sent (then send without reply_to_id).
            file_path: Optional local file to attach, as a link in the message that opens it in the browser.
                In a chat (1:1, group, meeting) it goes to YOUR OneDrive › "Microsoft Teams Chat Files", like
                Teams does (a same-name file is never replaced: "name (1).ext"). In a channel it goes to the
                SharePoint attachment folder (`sharepoint.attachment_folder`) as before - people with access to
                that site can open it.
            file_paths: Several local files to attach to this ONE message (e.g. ["/a/spec.pdf", "/a/plan.pdf"]) -
                no need to zip them. Combined with `file_path` (that one first), duplicates dropped. Uploaded in
                parallel; the message lists every file, in the order given. If any file is missing or fails to
                upload, NOTHING is sent and the error names the file(s).
            share_scope: Who can open a chat attachment: "members" (default) - only the chat's members, each
                granted by name, no invitation email; a member whose account cannot be looked up is left out
                and named in the result; "organization" - anyone in the company with the link.
                Ignored for channels.
            allow_org_fallback: Only for share_scope="members": when NO member can be looked up at all (or the
                member list cannot be read), give the file an organization link instead of refusing. Default
                false = refuse, nothing sent. Tell the user before setting it.
            mentions: People to tag: full name (diacritics optional; with "(Unit)" it must match
                exactly), email/UPN, alias ("hoangnh21") or MRI "8:orgid:<guid>". Chat history first,
                then the directory. Namesakes -> error listing candidates, nothing sent.
            timeout_seconds: Optional per-request timeout. Leave empty unless `file_path` is a large file; a slow send
                means a sick server - check whether it was sent before retrying.
        """
        if share_scope not in ("members", "organization"):
            raise Mcp365Error(f"share_scope phải là 'members' hoặc 'organization', không phải '{share_scope}'.")
        # With `file_paths`, every file must exist before a draft is shown (the draft lists sizes).
        # A lone `file_path` keeps the old flow exactly: checked when sending.
        attached = local_files(file_path, file_paths) if as_list(file_paths) else ([Path(file_path)] if file_path else [])
        conv = teams().find_conversation(chat_name_or_id)
        channel = conv.get("type") == "Channel" or "@thread.tacv2" in conv["id"]
        if channel and (dupes := same_names(attached)):
            raise Mcp365Error(
                f"Tin nhắn CHƯA được gửi: tên file trùng nhau ({', '.join(dupes)}); trong channel file cùng tên "
                "ghi đè nhau.",
                "Đổi tên cho khác nhau rồi gửi lại.",
            )
        people = teams().resolve_mentions(conv["id"], mentions) if mentions else []
        detail = message
        if attached:
            if channel:
                where = f"SharePoint › `{get_config().sharepoint.attachment_folder}`"
                who = "người có quyền trên site SharePoint đó"
            else:
                where = "OneDrive của bạn › `Microsoft Teams Chat Files`"
                who = ("chỉ các thành viên của chat này (cấp quyền từng người, không gửi email mời; ai không tra "
                       "được tài khoản thì bỏ qua và báo lại)"
                       if share_scope == "members" else "mọi người trong tổ chức có link")
                if share_scope == "members" and allow_org_fallback:
                    who += " - nếu không tra được thành viên nào thì cấp link cho mọi người trong tổ chức"
            if len(attached) == 1 and not as_list(file_paths):
                detail += f"\n\n**Đính kèm:** `{file_path}` → {where}\n**Ai mở được file:** {who}"
            else:
                rows = "\n".join(
                    f"{i}. `{p}` · {human_size(p.stat().st_size)}" for i, p in enumerate(attached, 1)
                )
                detail += (f"\n\n**Đính kèm ({len(attached)} file, cùng một tin)** → {where}:\n{rows}"
                           f"\n**Ai mở được file:** {who}")
        if people:
            detail += "\n\n**Tag:** " + ", ".join(mention_label(p) for p in people)
        approval.require_confirm(is_user_confirm, "Gửi tin nhắn Teams", f"{conv['name']} (`{conv['id']}`)", detail)
        return _render_send(
            teams().send_message(
                conversation_id_or_name=conv["id"],
                message=message,
                reply_to_id=reply_to_id or None,
                mentions=people,
                share_scope=share_scope,
                allow_org_fallback=allow_org_fallback,
                **_file_kwargs(file_path, file_paths, attached),
            )
        )

    def _render_send(res: dict[str, Any]) -> str:
        extra = []
        if res.get("reply_to_id"):
            extra.append(f"- **Trả lời tin nhắn:** `{res['reply_to_id']}`")
        files = res.get("attached_files") or ([res["attached_file"]] if res.get("attached_file") else [])
        rows = [f"[{f['name']}]({f.get('share_link') or f['webUrl']})"
                + (f" · lưu ở {f['location']}" if f.get("location") else "") for f in files]
        if len(rows) == 1:
            extra.append(f"- **File đính kèm:** {rows[0]}")
        elif rows:
            extra.append(f"- **File đính kèm ({len(rows)}):**\n" + "\n".join(f"  {i}. {r}" for i, r in enumerate(rows, 1)))
        if res.get("share_note"):
            extra.append(f"- ⚠️ **Quyền file:** {res['share_note']}")
        if res.get("share_skipped"):
            who = ", ".join(f"{p.get('name') or p['mri']} ({p['reason']})" for p in res["share_skipped"])
            extra.append(f"- ⚠️ **Không cấp quyền file được cho {len(res['share_skipped'])} thành viên** "
                         f"(không tra được tài khoản; họ sẽ không mở được file): {who}")
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
        message_id: str = "",
        include_quoted: bool = False,
        timeout_seconds: TransferTimeout = None,
        message_ids: list[str] | str | None = None,
    ) -> str:
        """Download files from a chat: paperclip attachments and SharePoint/OneDrive links.

        Attachments live in the message's ``properties.files`` (the HTML body of a
        file-only message is empty), so both sources are scanned, newest first.
        Files are downloaded in parallel; the report keeps their order.

        Args:
            chat_name_or_id: Chat name or thread ID.
            target_dir: Destination directory.
            limit: Maximum number of files to download.
            file_name: Optional case-insensitive substring to pick one file, e.g. "PSDK.zip".
            scan_messages: How many recent messages to scan.
            message_id: Optional: only this message (read by ID if it is older than `scan_messages`).
            include_quoted: Also download the files of the message(s) being quoted or forwarded (the original is
                read by ID when needed). A quote reply to a file-only message carries no file itself - use this.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
            message_ids: Several message IDs of this chat in one call (combined with `message_id`, that one
                first); their files come in the order the IDs are given. `limit` still caps the total number of
                files - raise it when the messages carry more.
        """
        client = teams()
        ids = as_list(message_id, message_ids)
        missing_ids: list[str] = []
        if len(ids) == 1:
            conv, msg = client.get_message(
                chat_name_or_id, ids[0], scan=scan_messages, resolve_quotes=include_quoted
            )
            chat_name, messages = conv["name"], [msg]
        elif ids:
            chat_name, messages, missing_ids = _messages_by_id(client, chat_name_or_id, ids, scan_messages, include_quoted)
        else:
            # Không truyền thêm tham số khi không cần: hành vi mặc định y như cũ.
            extra = {"resolve_quotes": True} if include_quoted else {}
            res = client.get_messages(chat_name_or_id, limit=scan_messages, **extra)
            chat_name, messages = res["conversation_name"], res["messages"]
        wanted = file_name.lower().strip()
        links: list[tuple[str, str]] = []
        origin: dict[tuple[str, str], str] = {}
        for msg in reversed(messages):
            candidates = [(a["name"], a["url"], "") for a in msg.get("attachments", [])]
            candidates += [(link.rsplit("/", 1)[-1], link, "") for link in msg.get("sharepoint_links", [])]
            candidates += [(img["name"], img["url"], "") for img in msg.get("images", [])]
            if include_quoted:
                for item in msg.get("quoted") or []:
                    tag = item.get("message_id") or "?"
                    candidates += [(a["name"], a["url"], tag) for a in item.get("attachments") or []]
                    candidates += [(img["name"], img["url"], tag) for img in item.get("images") or []]
            for name, link, quoted_from in candidates:
                if wanted and wanted not in urllib.parse.unquote(name).lower():
                    continue
                if (name, link) not in links:
                    links.append((name, link))
                    origin[(name, link)] = quoted_from
        missing_note = (
            f"\n\n> ⚠️ **Không đọc được {len(missing_ids)}/{len(ids)} tin:**\n" + "\n".join(f"> - {m}" for m in missing_ids)
            if missing_ids else ""
        )
        if not links:
            what = f"file khớp '{file_name}'" if wanted else "file đính kèm hay link SharePoint/OneDrive nào"
            where = ("tin " + ", ".join(f"`{i}`" for i in ids)) if ids else f"{scan_messages} tin gần nhất"
            hint = ""
            # Chỉ gợi ý khi hỏi đúng tin: lượt quét mặc định giữ nguyên câu trả lời cũ.
            quoted = [i for m in messages for i in m.get("quoted") or []] if ids else []
            if quoted and not include_quoted:
                qids = ", ".join(f"`{i['message_id']}`" for i in quoted if i.get("message_id")) or "không rõ id"
                hint = f"\n→ Tin này trích/chuyển tiếp tin {qids}: gọi lại với `include_quoted=true` để tải file của tin gốc."
            elif quoted:
                notes = [f"`{i['message_id']}`: {i['error']}" for i in quoted if i.get("error")]
                if notes:
                    hint = "\n→ Không đọc được tin gốc " + "; ".join(notes) + "."
            return f"Không tìm thấy {what} trong {where} của '{chat_name}'.{hint}{missing_note}"

        spc = sp() if any(not is_teams_media_url(link) for _n, link in links[:limit]) else None

        def fetch(entry: tuple[str, str]) -> str:
            name, link = entry
            if is_teams_media_url(link):
                dest = Path(target_dir or "downloads").expanduser().resolve()
                dest.mkdir(parents=True, exist_ok=True)
                out_p = client.download_image(link, dest / name)
                return f"✓ Đã tải ảnh: `{out_p}` ({human_size(out_p.stat().st_size)})"
            return spc.download_link(link, target_dir=target_dir)

        # Song song (http.transfer_workers); file cùng tên tải nối tiếp nhau, thứ tự báo cáo như cũ.
        picked = links[:limit]
        reports, failures = [], []
        for (name, link), (report, exc) in zip(
            picked, run_ordered(fetch, picked, key=lambda e: e[0].casefold()), strict=True
        ):
            prefix = f"↩️ File của tin được trích `{origin[(name, link)]}`:\n" if origin[(name, link)] else ""
            if exc is None:
                reports.append(prefix + report)
            elif isinstance(exc, Mcp365Error):
                failures.append(f"- {prefix.strip()} `{link[:70]}…`: {exc.message}")
            else:
                raise exc
        body = f"# Đã xử lý {len(reports)}/{min(len(links), limit)} tệp từ '{chat_name}'\n\n"
        body += "\n\n---\n\n".join(reports)
        if failures:
            body += "\n\n> ⚠️ **Thất bại:**\n" + "\n".join(f"> {f}" for f in failures)
        if len(links) > limit:
            body += f"\n\n> Còn {len(links) - limit} tệp nữa chưa tải (giới hạn `limit={limit}`): gọi lại với `limit` lớn hơn."
        return body + missing_note

    @mcp.tool()
    def download_message_images(
        chat_name_or_id: str,
        message_id: str = "",
        target_dir: str = "",
        limit: int = 5,
        include_quoted: bool = False,
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
            include_quoted: Also download the images of the message(s) being quoted or forwarded.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files, recordings or
                folders with many files.
        """
        extra = {"include_quoted": True} if include_quoted else {}
        downloaded = teams().download_message_images(
            chat_name_or_id, message_id=message_id, target_dir=target_dir, limit=limit, **extra
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
                + (f", được trích trong tin `{img['quoted_by']}`" if img.get("quoted_by") else "")
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


def _mail_targets(to: list[str], cc: list[str], bcc: list[str]) -> str:
    target = f"To: {', '.join(to) or '(trống)'}"
    if cc:
        target += f"\nCC: {', '.join(cc)}"
    if bcc:
        target += f"\nBCC: {', '.join(bcc)}"
    return target


def _mail_draft(
    subject: str,
    body: str,
    body_format: str,
    files: list[dict[str, Any]],
    link_scope: str,
    original: dict[str, Any] | None = None,
    mode: str = "",
    signature: dict[str, Any] | None = None,
) -> str:
    """Bản nháp cho người dùng duyệt: tiêu đề, nội dung, chữ ký, file kèm, mail gốc."""
    parts = [f"**Subject:** {subject}"]
    if original:
        label = {"reply": "Trả lời", "reply_all": "Trả lời tất cả", "forward": "Chuyển tiếp"}[mode]
        parts.append(
            f"**{label} mail gốc:** {original['subject']} — từ {original['from']} — {original['received']}\n"
            f"(Message ID `{original['id']}`; phần trích dẫn mail gốc giữ nguyên bên dưới nội dung)"
        )
    fmt = {"markdown": "markdown → HTML", "html": "HTML tự viết, gửi nguyên không lọc"}.get(body_format, "văn bản thường")
    parts.append(f"**Nội dung** ({fmt}):\n\n{body}")
    if signature:
        images = ", ".join(f"`{i['name']}`" for i in signature["images"])
        parts.append(f"**Chữ ký** (`{signature['path']}`{'; ảnh inline: ' + images if images else ''}):\n\n{signature['text']}")
    else:
        parts.append("**Chữ ký:** không gắn (signature=false).")
    if files:
        parts.append("**File kèm:**\n" + "\n".join(f"- {mail_mod.describe_attachment(f, link_scope)}" for f in files))
    if original and mode == "forward" and original.get("has_attachments"):
        parts.append("**Kèm theo:** các file đính kèm của mail gốc (Outlook tự chuyển tiếp cùng).")
    return "\n\n".join(parts)


def _mail_signature(include: bool) -> dict[str, Any] | None:
    """Chữ ký cho thư sắp gửi; bật mà chưa cấu hình thì báo lỗi, không gửi thư trơn."""
    if not include:
        return None
    sig = mail_mod.load_signature(get_config().mail.signature_dir)
    if sig is None:
        raise Mcp365Error(
            "Chưa cấu hình chữ ký mail, nên thư CHƯA được soạn.",
            "Đặt `[mail] signature_dir` trong ~/.config/mcp-auto-365-ms/config.toml (thư mục có signature.html + "
            "ảnh `cid:<tên file>`), hoặc gửi lại với signature=false nếu thư này không cần chữ ký.",
        )
    return sig


def _render_sent(result: dict[str, Any], what: str) -> str:
    lines = [f"✓ Outlook Web đã nhận {what} để gửi (HTTP {result['status']}; chưa phải xác nhận phát thành công)."]
    lines.append(f"- **To:** {', '.join(result['to'])}")
    if result["cc"]:
        lines.append(f"- **CC:** {', '.join(result['cc'])}")
    if result["bcc"]:
        lines.append(f"- **BCC:** {', '.join(result['bcc'])}")
    lines.append(f"- **Subject:** {result['subject']}")
    for f in result.get("attachments") or []:
        if f["method"] == "link":
            lines.append(f"- 📎 `{f['name']}` ({human_size(f['size'])}) · link: {f['share_link']} · lưu ở {f['location']}")
        else:
            how = "gắn trực tiếp (upload session)" if f["method"] == "session" else "gắn trực tiếp"
            lines.append(f"- 📎 `{f['name']}` ({human_size(f['size'])}) · {how}")
    if result.get("conversation_id"):
        lines.append(f"- **Conversation ID:** `{result['conversation_id']}`")
    lines.append("- **Lưu:** Sent Items")
    return "\n".join(lines)


def register_mail_tools(mcp) -> None:
    mcp = _ErrorAwareServer(mcp)

    @mcp.tool()
    def list_emails(
        folder: str = "inbox",
        limit: int = 20,
        unread_only: bool = False,
        since: str = "",
        query: str = "",
        sender: str = "",
        has_attachments: bool = False,
        timeout_seconds: RequestTimeout = None,
    ) -> str:
        """List recent or searched Outlook email from one mailbox folder.

        Args:
            folder: inbox, sent, drafts, deleted, archive, junk, a folder's display name (top level or under
                Inbox), or an Outlook folder ID.
            limit: Maximum messages to return (1-50).
            unread_only: Return only unread messages.
            since: Optional YYYY-MM-DD or ISO 8601 received-time lower bound.
            query: Optional Outlook mail search text (cannot be combined with the filters).
            sender: Optional sender. An email address filters exactly (combines with unread_only/since/
                has_attachments); a name becomes a `from:` search (cannot be combined with those filters).
            has_attachments: Return only messages with file attachments.
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
        """
        messages = outlook().list_messages(
            folder=folder,
            limit=limit,
            unread_only=unread_only,
            since=since,
            query=query,
            sender=sender,
            has_attachments=has_attachments,
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
        """Read one Outlook email in full using an ID returned by `list_emails`, with its attachment list.

        Args:
            message_id: Outlook message ID.
            timeout_seconds: Optional per-request timeout (default 30 s); raise only on a known-slow network.
        """
        msg = outlook().get_message(message_id)
        files = outlook().list_attachments(message_id)
        out = [
            f"# {msg['subject']}",
            f"- **Từ:** {msg['from']}",
            f"- **Tới:** {', '.join(msg['to']) or '(trống)'}",
        ]
        if msg["cc"]:
            out.append(f"- **CC:** {', '.join(msg['cc'])}")
        out.extend([f"- **Nhận:** {msg['received']}", f"- **Message ID:** `{msg['id']}`"])
        if files:
            out.append(f"- **Đính kèm ({len(files)}):**")
            for f in files:
                kind = {"item": " · mail đính kèm (.eml)", "reference": " · tệp đám mây (link)"}.get(f["kind"], "")
                inline = " · ảnh trong nội dung" if f["is_inline"] else ""
                out.append(f"  - `{f['name']}` · {human_size(f['size'])}{kind}{inline}")
            out.append("  → Tải về: `download_email_attachments(message_id)`.")
        else:
            out.append("- **Đính kèm:** Không")
        out.extend(["", "---", "", msg.get("body") or msg["preview"] or "(email không có nội dung)"])
        if msg["web_link"]:
            out.append(f"\n---\n[Mở trong Outlook]({msg['web_link']})")
        return "\n".join(out)

    @mcp.tool()
    def download_email_attachments(
        message_id: str = "",
        target_dir: str = "",
        file_name: str = "",
        include_inline: bool = False,
        timeout_seconds: TransferTimeout = None,
        message_ids: list[str] | str | None = None,
    ) -> str:
        """Download the attachments of one or more Outlook emails into a local folder.

        Files are never overwritten (a same-name file becomes "name (1).ext"). An attached email is saved
        as .eml; a cloud attachment (OneDrive/SharePoint link) is downloaded through SharePoint. A message's
        files are downloaded in parallel; the report keeps their order.

        Args:
            message_id: Outlook message ID (from `list_emails` / `read_email`).
            target_dir: Destination directory (default ./downloads).
            file_name: Optional case-insensitive substring to pick some files, e.g. "report.xlsx".
            include_inline: Also save images embedded in the body (default false).
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large files.
            message_ids: Several message IDs in one call (combined with `message_id`, that one first); the
                report has one section per email, and one email failing does not stop the others.
        """
        ids = as_list(message_id, message_ids)
        if not ids:
            raise Mcp365Error("Thiếu message_id.", "Truyền `message_id` (một mail) hoặc `message_ids` (nhiều mail).")
        if len(ids) > 1:
            return _download_many_emails(ids, target_dir, file_name, include_inline)
        res = outlook().download_attachments(
            ids[0], target_dir=target_dir, name_filter=file_name, include_inline=include_inline
        )
        return _render_email_downloads(res, file_name, include_inline)

    def _download_many_emails(ids: list[str], target_dir: str, file_name: str, include_inline: bool) -> str:
        sections, errors = [], []
        for mid in ids:  # file của mỗi mail đã tải song song bên trong
            try:
                res = outlook().download_attachments(
                    mid, target_dir=target_dir, name_filter=file_name, include_inline=include_inline
                )
            except Mcp365Error as exc:
                errors.append(f"`{mid}`: {exc.message}")
                continue
            sections.append(f"## Mail `{mid}`\n\n" + _render_email_downloads(res, file_name, include_inline))
        if not sections:
            raise Mcp365Error(
                f"Không tải được file của mail nào ({len(errors)} mail):\n" + "\n".join(f"- {e}" for e in errors),
                "Kiểm tra lại message ID (lấy từ `list_emails`).",
            )
        out = f"# File đính kèm của {len(ids)} mail\n\n" + "\n\n---\n\n".join(sections)
        if errors:
            out += f"\n\n> ⚠️ **Không đọc được {len(errors)}/{len(ids)} mail** (kết quả trên chưa đầy đủ):\n"
            out += "\n".join(f"> - {e}" for e in errors)
        return out

    def _render_email_downloads(res: dict[str, Any], file_name: str, include_inline: bool) -> str:
        if not res["matched"]:
            what = f"khớp '{file_name}'" if file_name.strip() else "nào"
            return f"Email này không có file đính kèm {what}" + ("" if include_inline else " (bỏ qua ảnh trong nội dung).")
        out = [f"# Đã tải {len(res['saved'])}/{res['matched']} file đính kèm vào `{res['target_dir']}`\n"]
        for f in res["saved"]:
            if f.get("report"):
                out.append(f"- `{f['name']}` (tệp đám mây):\n\n{f['report']}")
            else:
                out.append(f"- ✓ `{f['path']}` ({human_size(f['size'])})")
        if res["failures"]:
            out.append("\n> ⚠️ **Thất bại:**\n" + "\n".join(f"> - {x}" for x in res["failures"]))
        return "\n".join(out)

    @mcp.tool()
    def send_email(
        to: list[str],
        subject: str,
        body: str,
        is_user_confirm: approval.UserConfirm,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        attachments: list[str] | str | None = None,
        attach_mode: str = "auto",
        link_scope: str = "recipients",
        body_format: str = "markdown",
        signature: bool = True,
        timeout_seconds: TransferTimeout = None,
    ) -> str:
        """Send an Outlook email, optionally with files and simple formatting, after per-message human approval.

        ALWAYS show the user the exact To/CC/BCC, subject, body and attached files first. Call
        with false to receive that draft back; call with true only after the
        user explicitly approves this exact email. Approval for another message
        or a general instruction to send does not count.

        Args:
            to: Recipient email addresses.
            subject: Exact email subject.
            body: Exact email body.
            is_user_confirm: Required. True only after the user approved this exact email, recipients and files.
            cc: Optional CC recipient addresses.
            bcc: Optional BCC recipient addresses.
            attachments: Optional local file paths - a list (any number of files, all in this one email; no need
                to zip) or a single path string. Under 3 MB: attached directly; 3-150 MB: attached through an
                Outlook upload session; above 150 MB (or when the total would pass 150 MB): sent as a OneDrive link
                (link files are uploaded in parallel). Any missing file -> error listing every missing one, nothing
                sent; any upload failure -> nothing sent, the error names the file.
            attach_mode: "auto" (default, as above), "attach" (always attach; error if over the limit) or
                "link" (upload every file to YOUR OneDrive › Attachments and put links in the body).
            link_scope: Who can open linked files: "recipients" (default) - only the To/CC/BCC people, each
                granted by name, no invitation email; "organization" - anyone in the company with the link.
            body_format: "markdown" (default, like Teams messages: **bold**, *italic*, `code`, [text](url),
                "- " bullets, "1. " lists, tables, "#" headings, "> " quotes, ``` code blocks; sent as styled
                HTML), "html" (your own HTML, sent as is with no filtering - use inline `style=""`, Outlook
                drops `<style>` blocks; a full page is cut down to its `<body>`), or "text" (sent exactly as
                written, no character is interpreted). Bare http(s):// URLs are clickable in markdown and text
                whenever the body goes out as HTML.
            signature: Append your mail signature (default true; it comes from `[mail] signature_dir`, logo
                images attached inline). Set false only when this mail must go without it; with no signature
                configured, true is an error instead of a silent unsigned mail.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large attachments.
        """
        to = mail_mod.clean_addresses(to)
        cc = mail_mod.clean_addresses(cc)
        bcc = mail_mod.clean_addresses(bcc)
        mail_mod.render_body_html("", body_format)  # sai body_format thì báo ngay, trước bản nháp
        if link_scope not in mail_mod.LINK_SCOPES:
            raise Mcp365Error(f"link_scope phải là 'recipients' hoặc 'organization', không phải '{link_scope}'.")
        files = mail_mod.plan_attachments(attachments, attach_mode)
        sig = _mail_signature(signature)
        approval.require_confirm(
            is_user_confirm,
            "Gửi email Outlook",
            _mail_targets(to, cc, bcc),
            _mail_draft(subject, body, body_format, files, link_scope, signature=sig),
        )
        result = outlook().compose_and_send(
            to=to, subject=subject, body=body, cc=cc, bcc=bcc, body_format=body_format,
            attachments=files, link_scope=link_scope, signature=sig,
        )
        return _render_sent(result, "email")

    @mcp.tool()
    def reply_email(
        message_id: str,
        body: str,
        is_user_confirm: approval.UserConfirm,
        mode: str = "reply",
        to: list[str] | None = None,
        cc: list[str] | None = None,
        bcc: list[str] | None = None,
        attachments: list[str] | str | None = None,
        attach_mode: str = "auto",
        link_scope: str = "recipients",
        body_format: str = "markdown",
        signature: bool = True,
        timeout_seconds: TransferTimeout = None,
    ) -> str:
        """Reply, reply-all or forward an Outlook email in the same thread, after per-message human approval.

        The original is quoted below your text, like Outlook does. ALWAYS show the user the draft first:
        call with false to get it (the exact To/CC/BCC, subject, body, files and the original mail);
        call with true only after the user explicitly approves this exact reply.

        Args:
            message_id: ID of the email to answer (from `list_emails` / `read_email`).
            body: Exact text to put above the quoted original.
            is_user_confirm: Required. True only after the user approved this exact reply, recipients and files.
            mode: "reply" (to the sender), "reply_all" (sender + all To/CC, minus you) or "forward" (needs `to`;
                the original's attachments go along).
            to: Extra To addresses (required for forward).
            cc: Extra CC addresses.
            bcc: Optional BCC addresses.
            attachments: Optional local file paths (a list or one path string), handled as in `send_email`.
            attach_mode: "auto" (default), "attach" or "link" - as in `send_email`.
            link_scope: "recipients" (default) or "organization" - as in `send_email`.
            body_format: "markdown" (default), "html" (raw, unfiltered) or "text" - as in `send_email`.
            signature: Append your signature above the quoted original (default true) - as in `send_email`.
            timeout_seconds: Optional per-request timeout (default 120 s). Raise it for large attachments.
        """
        mail_mod.render_body_html("", body_format)
        if link_scope not in mail_mod.LINK_SCOPES:
            raise Mcp365Error(f"link_scope phải là 'recipients' hoặc 'organization', không phải '{link_scope}'.")
        files = mail_mod.plan_attachments(attachments, attach_mode)
        sig = _mail_signature(signature)
        plan = outlook().prepare_reply(message_id, mode, to=to, cc=cc, bcc=bcc)
        label = {"reply": "Trả lời email", "reply_all": "Trả lời tất cả email", "forward": "Chuyển tiếp email"}[mode]
        approval.require_confirm(
            is_user_confirm,
            label,
            _mail_targets(plan["to"], plan["cc"], plan["bcc"]),
            _mail_draft(plan["subject"], body, body_format, files, link_scope, plan["original"], mode, signature=sig),
        )
        result = outlook().compose_and_send(
            to=plan["to"], cc=plan["cc"], bcc=plan["bcc"], subject=plan["subject"], body=body,
            body_format=body_format, attachments=files, link_scope=link_scope,
            reply_to_id=plan["original"]["id"], mode=mode, signature=sig,
        )
        return _render_sent(result, label[0].lower() + label[1:])


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
