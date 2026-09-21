"""In-memory workbook reading and editing: no network involved."""

from __future__ import annotations

import io

import pytest
from openpyxl import Workbook, load_workbook

from common.errors import Mcp365Error
from sharepoint import sheets


def _checklist() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "SYS2_PIN"
    ws.append(["Feature", "Battery"])
    ws.append(["ID", "Review Item", "Description", "Yes/No", "Comment"])
    ws.append(["SYS2-001", "Complete Description", "Mô tả đầy đủ", "Yes", ""])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def test_render_lists_sheets_when_none_is_named():
    out = sheets.render_sheet(_checklist(), name="c.xlsx")
    assert "`SYS2_PIN`" in out


def test_render_shows_a1_addresses():
    out = sheets.render_sheet(_checklist(), sheet="SYS2_PIN", name="c.xlsx")
    assert "| # | A | B | C | D | E |" in out
    assert "| 3 | SYS2-001 |" in out


def test_render_unknown_sheet_lists_the_real_ones():
    with pytest.raises(Mcp365Error) as excinfo:
        sheets.render_sheet(_checklist(), sheet="nope")
    assert "SYS2_PIN" in excinfo.value.remediation


def test_apply_edits_existing_sheet_and_logs_old_values():
    data, changes = sheets.apply_cells(_checklist(), "SYS2_PIN", {"d3": "No", "E3": "Thiếu catalog"})
    ws = load_workbook(io.BytesIO(data))["SYS2_PIN"]
    assert ws["D3"].value == "No"
    assert ws["E3"].value == "Thiếu catalog"
    assert {"cell": "D3", "old": "Yes", "new": "No"} in changes


def test_missing_sheet_is_cloned_from_a_template():
    """A new feature gets the same checklist rows as its siblings."""
    data, changes = sheets.apply_cells(_checklist(), "SYS2_INSPECTION", {"B1": "Inspection"}, copy_sheet_from="SYS2_PIN")
    wb = load_workbook(io.BytesIO(data))
    assert wb["SYS2_INSPECTION"]["A3"].value == "SYS2-001"
    assert wb["SYS2_INSPECTION"]["B1"].value == "Inspection"
    assert wb["SYS2_PIN"]["B1"].value == "Battery"  # template untouched
    assert changes[0]["cell"] == "(sheet)"


def test_bad_template_name_is_actionable():
    with pytest.raises(Mcp365Error):
        sheets.apply_cells(_checklist(), "NEW", {"A1": "x"}, copy_sheet_from="MISSING")


def test_bad_cell_address_is_rejected():
    with pytest.raises(Mcp365Error):
        sheets.apply_cells(_checklist(), "SYS2_PIN", {"not-a-cell": "x"})


def test_non_xlsx_bytes_explain_the_likely_cause():
    """A 'downloaded' xlsx is sometimes an HTML login page - say so."""
    with pytest.raises(Mcp365Error) as excinfo:
        sheets.render_sheet(b"<html>sign in</html>", name="x.xlsx")
    assert "HTML" in excinfo.value.remediation


def test_http_412_becomes_a_concurrent_edit_error():
    import urllib.error

    from common.errors import ConcurrentEditError, classify_http_error

    exc = urllib.error.HTTPError("https://graph", 412, "Precondition Failed", {}, io.BytesIO(b"{}"))
    assert isinstance(classify_http_error(exc, "ghi"), ConcurrentEditError)
