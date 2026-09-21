"""Read and edit ``.xlsx`` workbooks in memory.

Pure functions over bytes: no network, no credentials. The SharePoint client
fetches and uploads; this module only turns bytes into a table or applies a set
of cell edits and reports what changed.

Round-tripping through openpyxl keeps values, styles, merged cells and cell
comments, but **drops charts and images**. That is acceptable for review
checklists and trackers, not for a dashboard workbook.
"""

from __future__ import annotations

import io
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils.cell import coordinate_from_string
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


def _cell_text(value: Any) -> str:
    if value is None:
        return ""
    return str(value).replace("|", "\\|").replace("\n", " ")


def render_sheet(data: bytes, sheet: str = "", max_rows: int = 60, name: str = "") -> str:
    """List the sheets, or render one sheet as a Markdown table with A1 refs."""
    wb = _open(data, name)
    if not sheet:
        rows = [f"# Workbook `{name}` — {len(wb.sheetnames)} sheet\n", "| Sheet | Kích thước |", "| --- | --- |"]
        rows += [f"| `{ws.title}` | {ws.max_row} × {ws.max_column} |" for ws in wb.worksheets]
        return "\n".join(rows)

    if sheet not in wb.sheetnames:
        raise Mcp365Error(f"Không có sheet '{sheet}'.", f"Các sheet hiện có: {', '.join(wb.sheetnames)}")
    ws = wb[sheet]
    width = ws.max_column
    letters = [ws.cell(row=1, column=c).column_letter for c in range(1, width + 1)]
    out = [f"# `{name}` › `{sheet}` ({ws.max_row} × {width})\n", "| # | " + " | ".join(letters) + " |"]
    out.append("| --- " * (width + 1) + "|")
    for r, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, max_rows), values_only=True), start=1):
        out.append(f"| {r} | " + " | ".join(_cell_text(v) for v in row) + " |")
    if ws.max_row > max_rows:
        out.append(f"\n_… còn {ws.max_row - max_rows} dòng nữa (tăng `max_rows` để xem)._")
    return "\n".join(out)


def apply_cells(
    data: bytes, sheet: str, cells: dict[str, Any], copy_sheet_from: str = "", name: str = ""
) -> tuple[bytes, list[dict[str, Any]]]:
    """Write ``cells`` (``{"E3": "No"}``) into ``sheet`` and return the new bytes.

    A missing sheet is created, optionally as a copy of ``copy_sheet_from`` so a
    new feature gets the same checklist rows and formatting as its siblings.
    Returns the change log so the caller can show exactly what will be written.
    """
    if not cells and not copy_sheet_from:
        raise Mcp365Error("Không có ô nào để ghi.", "Truyền `cells`, VD {\"E3\": \"No\"}.")

    wb = _open(data, name)
    changes: list[dict[str, Any]] = []

    if sheet in wb.sheetnames:
        ws = wb[sheet]
    elif copy_sheet_from:
        if copy_sheet_from not in wb.sheetnames:
            raise Mcp365Error(
                f"Không có sheet mẫu '{copy_sheet_from}'.", f"Các sheet hiện có: {', '.join(wb.sheetnames)}"
            )
        ws = wb.copy_worksheet(wb[copy_sheet_from])
        ws.title = sheet
        changes.append({"cell": "(sheet)", "old": "", "new": f"tạo mới, sao từ '{copy_sheet_from}'"})
    else:
        ws = wb.create_sheet(sheet)
        changes.append({"cell": "(sheet)", "old": "", "new": "tạo mới (trống)"})

    for ref, value in cells.items():
        ref = ref.strip().upper()
        try:
            coordinate_from_string(ref)
        except (CellCoordinatesException, ValueError) as exc:
            raise Mcp365Error(f"Địa chỉ ô không hợp lệ: '{ref}'.", "Dùng dạng A1, VD 'E3'.") from exc
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
