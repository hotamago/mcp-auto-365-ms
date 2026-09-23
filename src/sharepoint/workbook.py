"""Per-cell Excel edits through the Microsoft Graph workbook API.

**Why this exists.** ``update_sharepoint_sheet`` used to download the whole
workbook, edit it with openpyxl and PUT the whole file back with ``If-Match``.
That is a document-level write, so it loses to anyone who has the file open: on
23/09 a single-cell edit to ``ViTa - S5 - Management Plan.xlsx`` was refused
three times in a row (423 locked, 412 eTag changed, 423) while the same person
was editing the same sheet in Excel Online without trouble. The browser wins
because it edits *cells*, not files.

The Graph workbook API is that same cell-level channel::

    PATCH /drives/{drive}/items/{item}/workbook/worksheets/{id}/range(address='Q34')
    {"values": [["..."]]}

It co-authors instead of overwriting: it works while others have the file open,
it cannot silently drop a concurrent change to another cell, and it never
round-trips the file through openpyxl, so charts and images survive.

**Channel.** Graph only, with the Azure CLI token. SharePoint's own embedded
``/_api/v2.0`` endpoint - the primary channel for every other operation in this
package - does not implement ``/workbook`` and answers ``404 itemNotFound``, so
these calls deliberately skip ``call_sharepoint_or_graph`` and go straight to
Graph.

**Merged cells.** Graph v1.0 exposes no merged-area query, so a PATCH into a hidden non-anchor
cell of a merged range is accepted and never displayed. ``sheets.render_sheet`` lists a sheet's
merged ranges under its table; write to the anchor address it names.

**Fallback.** Anything this module cannot do raises :class:`WorkbookUnsupported`
*before* the first cell is written, and the caller falls back to the old
download-edit-upload path, saying which channel it used and why. Once a cell has
been written there is no fallback: re-uploading the whole file would undo the
co-authoring the write just bought.
"""

from __future__ import annotations

import re
import urllib.parse
from typing import Any, Protocol

from common.errors import Mcp365Error

#: A single A1 cell. Ranges ("A1:B2") and whole columns are rejected: the change
#: log and the readback are per cell, and a range hides what it overwrites.
_CELL_RE = re.compile(r"^[A-Z]{1,3}[1-9][0-9]{0,6}$")

#: Only Office Open XML workbooks. The Excel REST API does not serve .xls, and
#: .xlsm macro workbooks are not supported either.
SUPPORTED_SUFFIX = ".xlsx"


class GraphCall(Protocol):
    """How this module reaches Graph. Injected so tests never touch the network."""

    def __call__(
        self,
        path: str,
        method: str = "GET",
        body: dict | None = None,
        session_id: str = "",
        context: str = "",
    ) -> dict[str, Any]: ...


class WorkbookUnsupported(Exception):
    """The workbook API cannot serve this edit - fall back, do not fail.

    Carries a human-readable ``reason`` the caller reports next to the fallback,
    so the downgrade is never silent.
    """

    def __init__(self, reason: str) -> None:
        super().__init__(reason)
        self.reason = reason


def _cell_text(value: Any) -> str:
    if value is None or value == "":
        return ""
    return str(value).replace("|", "\\|").replace("\n", " ")


def check_addresses(cells: dict[str, Any]) -> dict[str, Any]:
    """Normalise ``{"q34": v}`` to ``{"Q34": v}``; reject anything but one cell."""
    cleaned: dict[str, Any] = {}
    for ref, value in cells.items():
        norm = str(ref).strip().upper().replace("$", "")
        if not _CELL_RE.match(norm):
            raise Mcp365Error(
                f"Địa chỉ ô không hợp lệ: '{ref}'.",
                "Dùng đúng một ô dạng A1, VD 'Q34'. Vùng ('A1:B2') hay cả cột ('A:A') không nhận.",
            )
        cleaned[norm] = value
    return cleaned


def workbook_base(drive_id: str, item_id: str) -> str:
    return f"/drives/{drive_id}/items/{item_id}/workbook"


def _sheet_segment(sheet_id: str) -> str:
    """The URL segment for a worksheet id.

    A worksheet id is a braced GUID (``{0E67021A-C793-4C5A-81DE-045870E88EAE}``)
    and goes in **bare**, percent-encoded: Graph answers ``404 ItemNotFound`` for
    the quoted key form ``worksheets('{...}')`` that works for sheet *names*.
    """
    return urllib.parse.quote(sheet_id, safe="")


def _single(range_payload: dict[str, Any]) -> Any:
    """The one value of a 1x1 ``workbookRange`` payload."""
    values = range_payload.get("values")
    if isinstance(values, list) and values and isinstance(values[0], list) and values[0]:
        return values[0][0]
    return None


def _worksheets(graph: GraphCall, base: str) -> list[dict[str, Any]]:
    """List the sheets. A failure here means the workbook API is unusable."""
    try:
        res = graph(f"{base}/worksheets?$select=id,name", context="liệt kê sheet qua Graph workbook API")
    except Mcp365Error as exc:
        raise WorkbookUnsupported(f"Graph workbook API không dùng được: {exc}") from exc
    sheets = res.get("value")
    if not isinstance(sheets, list):
        raise WorkbookUnsupported("Graph workbook API không trả danh sách sheet.")
    return sheets


def _sheet_id(sheets: list[dict[str, Any]], sheet: str) -> str:
    """Resolve a sheet name to its stable id.

    The id goes in the URL instead of the name: sheet names carry spaces,
    parentheses, Vietnamese diacritics and apostrophes (``Sprint 4 (17.09)``),
    and each of those would need its own quoting inside an OData function call.
    """
    for entry in sheets:
        if entry.get("name") == sheet:
            return str(entry.get("id") or "")
    return ""


def plan(
    graph: GraphCall,
    drive_id: str,
    item_id: str,
    sheet: str,
    cells: dict[str, Any],
    file_name: str = "",
    copy_sheet_from: str = "",
) -> tuple[str, list[dict[str, Any]]]:
    """Read what each cell holds today and return ``(sheet_id, change log)``.

    Nothing is written and no write session is opened, so this is what the
    unconfirmed (preview) call runs. Raises :class:`WorkbookUnsupported` when the
    edit needs the whole-file path instead.
    """
    if not file_name.lower().endswith(SUPPORTED_SUFFIX):
        raise WorkbookUnsupported(
            f"Graph workbook API chỉ hỗ trợ `{SUPPORTED_SUFFIX}`, file này là `{file_name or '(không rõ)'}`."
        )
    if copy_sheet_from:
        raise WorkbookUnsupported(
            "Graph workbook API không nhân bản được sheet kèm định dạng (`copy_sheet_from`)."
        )
    if not cells:
        raise Mcp365Error("Không có ô nào để ghi.", 'Truyền `cells`, VD {"Q34": "Đã xong"}.')

    cells = check_addresses(cells)
    base = workbook_base(drive_id, item_id)
    sheet_id = _sheet_id(_worksheets(graph, base), sheet)

    changes: list[dict[str, Any]] = []
    if not sheet_id:
        # A brand-new empty sheet is one call; cloning one with its formatting is
        # not, which is why copy_sheet_from falls back above.
        changes.append({"cell": "(sheet)", "old": "", "new": f"tạo mới (trống) sheet '{sheet}'"})
        changes += [{"cell": ref, "old": "", "new": _cell_text(v)} for ref, v in cells.items()]
        return "", changes

    for ref, value in cells.items():
        try:
            current = graph(
                f"{base}/worksheets/{_sheet_segment(sheet_id)}/range(address='{ref}')?$select=values",
                context=f"đọc ô {ref} qua Graph workbook API",
            )
        except Mcp365Error as exc:
            raise WorkbookUnsupported(f"Không đọc được ô {ref} qua Graph workbook API: {exc}") from exc
        changes.append({"cell": ref, "old": _cell_text(_single(current)), "new": _cell_text(value)})
    return sheet_id, changes


def _open_session(graph: GraphCall, base: str, warnings: list[str]) -> str:
    try:
        res = graph(
            f"{base}/createSession",
            method="POST",
            body={"persistChanges": True},
            context="mở phiên ghi workbook",
        )
        return str(res.get("id") or "")
    except Mcp365Error as exc:
        warnings.append(f"Không mở được phiên workbook ({exc}); ghi từng ô riêng lẻ.")
        return ""


def _close_session(graph: GraphCall, base: str, session_id: str, warnings: list[str]) -> None:
    if not session_id:
        return
    try:
        graph(f"{base}/closeSession", method="POST", body={}, session_id=session_id, context="đóng phiên workbook")
    except Mcp365Error as exc:
        # The session expires on its own; a failed close never fails the write.
        warnings.append(f"Không đóng được phiên workbook ({exc}); phiên sẽ tự hết hạn.")


def apply(
    graph: GraphCall,
    drive_id: str,
    item_id: str,
    sheet: str,
    sheet_id: str,
    cells: dict[str, Any],
) -> tuple[int, list[str]]:
    """Write the cells one PATCH at a time. Returns ``(written, warnings)``.

    A persistent session batches the writes; if the service refuses to open one
    the writes still go through sessionless, each in its own implicit session -
    slower, same result. The session is always closed.

    After the first successful PATCH there is no going back to the whole-file
    path: a later failure is reported as a partial write listing what landed,
    never retried by re-uploading the workbook.
    """
    cells = check_addresses(cells)
    base = workbook_base(drive_id, item_id)
    warnings: list[str] = []

    if not sheet_id:
        added = graph(
            f"{base}/worksheets/add", method="POST", body={"name": sheet}, context=f"tạo sheet '{sheet}'"
        )
        sheet_id = str(added.get("id") or "")
        if not sheet_id:
            raise Mcp365Error(
                f"Graph không trả id cho sheet mới '{sheet}'.",
                "Tạo sheet bằng Excel Online rồi chạy lại.",
            )

    session_id = _open_session(graph, base, warnings)
    written = 0
    try:
        for ref, value in cells.items():
            res = graph(
                f"{base}/worksheets/{_sheet_segment(sheet_id)}/range(address='{ref}')",
                method="PATCH",
                body={"values": [[value]]},
                session_id=session_id,
                context=f"ghi ô {ref} qua Graph workbook API",
            )
            written += 1
            # The PATCH answers with the updated range, so the readback is free.
            echoed = _single(res)
            if echoed is not None and _cell_text(echoed) != _cell_text(value):
                warnings.append(
                    f"Ô `{ref}`: Excel lưu thành `{_cell_text(echoed)}` (đã gửi `{_cell_text(value)}`)."
                )
    except Mcp365Error as exc:
        if written:
            raise Mcp365Error(
                f"Đã ghi {written}/{len(cells)} ô rồi mới gặp lỗi: {exc}",
                "Các ô đã ghi vẫn nằm trên file. Chạy lại chỉ với những ô còn thiếu.",
            ) from exc
        raise
    finally:
        _close_session(graph, base, session_id, warnings)
    return written, warnings


def render_changes(changes: list[dict[str, Any]], sheet: str) -> str:
    rows = [f"Sheet `{sheet}` — {len(changes)} thay đổi:\n", "| Ô | Hiện tại | Sẽ ghi |", "| --- | --- | --- |"]
    rows += [f"| `{c['cell']}` | {c['old'] or '_(trống)_'} | {c['new']} |" for c in changes]
    return "\n".join(rows)
