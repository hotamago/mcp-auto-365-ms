"""Read and edit ``.xlsx`` workbooks in memory.

Pure functions over bytes: no network, no credentials. The SharePoint client
fetches and uploads; this module only turns bytes into a table or applies a set
of cell edits and reports what changed.

Round-tripping through openpyxl keeps values, styles, merged cells and cell
comments, but **drops charts and images**. That is acceptable for review
checklists and trackers, not for a dashboard workbook.

**Merged cells.** A merged range stores its value in the top-left (anchor) cell
only; every other cell of the range is an openpyxl ``MergedCell`` with a
read-only ``None`` value. Both functions here follow that model rather than
spreading the anchor's value: ``render_sheet`` prints the anchor's value at the
anchor address and leaves the rest blank, and ``apply_cells`` refuses a write to
a non-anchor address instead of silently doing nothing. That keeps every A1
address printed by ``render_sheet`` usable as-is for ``update_sharepoint_sheet``.
"""

from __future__ import annotations

import io
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.cell import column_index_from_string, coordinate_from_string
from openpyxl.utils.exceptions import CellCoordinatesException

from common.errors import Mcp365Error


def _open(data: bytes, name: str = ""):
    try:
        return load_workbook(io.BytesIO(data), keep_vba=name.lower().endswith(".xlsm"))
    except Exception as exc:
        raise Mcp365Error(
            f"Không mở được file Excel{f' `{name}`' if name else ''}: {exc}",
            "File có thể không phải .xlsx/.xlsm thật (VD: bản tải về là trang HTML đăng nhập). "
            "Tải lại hoặc mở bằng Excel Online để kiểm tra.",
        ) from exc


# How many merged ranges the footer lists before summarising the rest.
_MERGED_PREVIEW = 12


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("|", "\\|").replace("\n", " ")


def _merged_ranges(ws) -> list[Any]:
    """The sheet's merged ranges, top-left first, as a stable list."""
    return sorted(ws.merged_cells.ranges, key=lambda r: (r.min_row, r.min_col))


def _merged_anchor(ranges: list[Any], row: int, col: int) -> str:
    """The anchor address if ``(row, col)`` is a *non-anchor* merged cell.

    Ranges are scanned rather than expanded into a per-cell map: a single
    whole-row merge (``A1:XFD1``) would otherwise allocate 16k entries.
    """
    for rng in ranges:
        if rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col:
            if (row, col) == (rng.min_row, rng.min_col):
                return ""
            return f"{get_column_letter(rng.min_col)}{rng.min_row}"
    return ""


def find_sheet(names: list[str], sheet: str) -> str:
    """The workbook's own name for ``sheet``: exact match first, then ignoring case
    and surrounding spaces (real names carry trailing spaces: ``'S5 Feature Release Plan '``).
    Returns ``""`` when there is none."""
    if sheet in names:
        return sheet
    wanted = sheet.casefold().strip()
    return next((n for n in names if n.casefold().strip() == wanted), "")


def merged_anchor_for(ranges: list[Any], ref: str) -> str:
    """The anchor to write to instead of ``ref``, or ``""`` if ``ref`` is writable."""
    try:
        col_letter, row_idx = coordinate_from_string(ref.strip().upper())
    except (CellCoordinatesException, ValueError) as exc:
        raise Mcp365Error(f"Địa chỉ ô không hợp lệ: '{ref}'.", "Dùng dạng A1, VD 'E3'.") from exc
    return _merged_anchor(ranges, row_idx, column_index_from_string(col_letter))


def merged_cell_error(ref: str, anchor: str) -> Mcp365Error:
    return Mcp365Error(
        f"Ô `{ref}` nằm trong một vùng gộp, không ghi trực tiếp được.",
        f"Ghi vào ô góc trên trái của vùng gộp: `{anchor}`.",
    )


def check_merged(data: bytes, sheet: str, refs: list[str], name: str = "") -> None:
    """Refuse any of ``refs`` that is a non-anchor cell of a merged range in ``sheet``.

    Used by the per-cell Graph path, which has no merged-area query of its own.
    A sheet missing from ``data`` has no merges to guard.
    """
    wb = _open(data, name)
    actual = find_sheet(wb.sheetnames, sheet)
    if not actual:
        return
    merged = _merged_ranges(wb[actual])
    for ref in refs:
        anchor = merged_anchor_for(merged, ref)
        if anchor:
            raise merged_cell_error(ref, anchor)


def render_sheet(data: bytes, sheet: str = "", max_rows: int = 60, name: str = "") -> str:
    """List the sheets, or render one sheet as a Markdown table with A1 refs.

    Merged ranges keep openpyxl's model: only the top-left (anchor) cell carries
    the value, the rest of the range prints blank. The ranges are listed under
    the table so the anchor to write to is visible - ``update_sharepoint_sheet``
    only accepts anchor addresses.
    """
    wb = _open(data, name)
    if not sheet:
        rows = [f"# Workbook `{name}` — {len(wb.sheetnames)} sheet\n", "| Sheet | Kích thước |", "| --- | --- |"]
        rows += [f"| `{ws.title}` | {ws.max_row} × {ws.max_column} |" for ws in wb.worksheets]
        return "\n".join(rows)

    actual = find_sheet(wb.sheetnames, sheet)
    if not actual:
        raise Mcp365Error(f"Không có sheet '{sheet}'.", f"Các sheet hiện có: {', '.join(wb.sheetnames)}")
    ws = wb[actual]
    width = ws.max_column
    # ``cell().column_letter`` blows up on a MergedCell (it has no address of its
    # own), so the header is built from the column index instead.
    letters = [get_column_letter(c) for c in range(1, width + 1)]
    out = [f"# `{name}` › `{sheet}` ({ws.max_row} × {width})\n", "| # | " + " | ".join(letters) + " |"]
    out.append("| --- " * (width + 1) + "|")
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), values_only=True), start=1):
        out.append(f"| {r} | " + " | ".join(_cell_text(v) for v in row) + " |")
    if ws.max_row > max_rows:
        out.append(f"\n_… còn {ws.max_row - max_rows} dòng nữa (tăng `max_rows` để xem)._")
    ranges = _merged_ranges(ws)
    if ranges:
        shown = ", ".join(f"`{r}`" for r in ranges[:_MERGED_PREVIEW])
        more = f" … và {len(ranges) - _MERGED_PREVIEW} vùng nữa" if len(ranges) > _MERGED_PREVIEW else ""
        out.append(
            f"\n_Ô gộp ({len(ranges)}): {shown}{more}. Giá trị nằm ở ô góc trên trái, "
            "các ô còn lại để trống — ghi vào ô góc trên trái._"
        )
    return "\n".join(out)


def apply_cells(
    data: bytes, sheet: str, cells: dict[str, Any], copy_sheet_from: str = "", name: str = ""
) -> tuple[bytes, list[dict[str, Any]]]:
    """Write ``cells`` (``{"E3": "No"}``) into ``sheet`` and return the new bytes.

    A missing sheet is created, optionally as a copy of ``copy_sheet_from`` so a
    new feature gets the same checklist rows and formatting as its siblings.
    Returns the change log so the caller can show exactly what will be written.

    An address inside a merged range but not its top-left cell is refused: that
    cell is read-only in openpyxl and invisible in Excel, so writing it would
    look like a silent no-op. The error names the anchor to use instead.
    """
    if not cells and not copy_sheet_from:
        raise Mcp365Error("Không có ô nào để ghi.", "Truyền `cells`, VD {\"E3\": \"No\"}.")

    wb = _open(data, name)
    changes: list[dict[str, Any]] = []

    existing = find_sheet(wb.sheetnames, sheet)
    if existing:
        ws = wb[existing]
    elif copy_sheet_from:
        template = find_sheet(wb.sheetnames, copy_sheet_from)
        if not template:
            raise Mcp365Error(
                f"Không có sheet mẫu '{copy_sheet_from}'.", f"Các sheet hiện có: {', '.join(wb.sheetnames)}"
            )
        ws = wb.copy_worksheet(wb[template])
        ws.title = sheet
        changes.append({"cell": "(sheet)", "old": "", "new": f"tạo mới, sao từ '{copy_sheet_from}'"})
    else:
        ws = wb.create_sheet(sheet)
        changes.append({"cell": "(sheet)", "old": "", "new": "tạo mới (trống)"})

    merged = _merged_ranges(ws)
    for ref, value in cells.items():
        ref = ref.strip().upper()
        anchor = merged_anchor_for(merged, ref)
        if anchor:
            raise merged_cell_error(ref, anchor)
        old = ws[ref].value
        ws[ref] = value
        changes.append({"cell": ref, "old": _cell_text(old), "new": _cell_text(value)})

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), changes


def render_changes(changes: list[dict[str, Any]], sheet: str) -> str:
    rows = [f"Sheet `{sheet}` — {len(changes)} thay đổi:\n", "| Ô | Hiện tại | Sẽ ghi |", "| --- | --- | --- |"]
    rows += [f"| `{c['cell']}` | {c['old'] or '_(trống)_'} | {c['new']} |" for c in changes]
    return "\n".join(rows)
