"""Per-cell Excel writes through the Graph workbook API - request shapes offline.

The live PATCH is not exercised here (tests never touch the network), so these
tests pin down exactly what would go on the wire: which URL, which body, which
headers, in which order, and what happens when any of it fails.
"""

from __future__ import annotations

import pytest

from common.errors import Mcp365Error
from sharepoint import workbook

DRIVE = "b!drive"
ITEM = "01ITEM"
BASE = f"/drives/{DRIVE}/items/{ITEM}/workbook"
SHEET = "Sprint 4 (17.09)"
SHEET_ID = "{0E67021A-C793-4C5A-81DE-045870E88EAE}"
ENCODED_ID = "%7B0E67021A-C793-4C5A-81DE-045870E88EAE%7D"


class FakeGraph:
    """Records every call and replays canned answers."""

    def __init__(self, sheets=((SHEET, SHEET_ID),), fail=None, values=None):
        self.calls: list[dict] = []
        self._sheets = sheets
        self._fail = fail or {}
        self._values = values or {}

    def __call__(self, path, method="GET", body=None, session_id="", context=""):
        self.calls.append(
            {"path": path, "method": method, "body": body, "session_id": session_id, "context": context}
        )
        for marker, exc in self._fail.items():
            if marker in path and (method != "GET" or marker.startswith("range")):
                raise exc
            if marker in path and method == "GET":
                raise exc
        if "/worksheets?" in path:
            return {"value": [{"name": n, "id": i} for n, i in self._sheets]}
        if path.endswith("/createSession"):
            return {"id": "SESSION-1", "persistChanges": True}
        if path.endswith("/closeSession"):
            return {}
        if path.endswith("/worksheets/add"):
            return {"id": "{NEW-SHEET}", "name": body.get("name")}
        if "/range(address=" in path:
            ref = path.split("address='")[1].split("'")[0]
            if method == "PATCH":
                return {"address": f"'{SHEET}'!{ref}", "values": body["values"]}
            return {"address": f"'{SHEET}'!{ref}", "values": [[self._values.get(ref, "")]]}
        raise AssertionError(f"unexpected call {method} {path}")

    def paths(self, method=None):
        return [c["path"] for c in self.calls if method is None or c["method"] == method]


# ------------------------------------------------------------------ addresses


def test_addresses_are_normalised():
    assert workbook.check_addresses({" q34 ": "a", "$B$7": "b"}) == {"Q34": "a", "B7": "b"}


@pytest.mark.parametrize("bad", ["A1:B2", "A:A", "1", "not-a-cell", "A0", ""])
def test_only_single_cells_are_accepted(bad):
    with pytest.raises(Mcp365Error):
        workbook.check_addresses({bad: "x"})


# ----------------------------------------------------------------- fallbacks


def test_xlsm_falls_back_to_the_whole_file_path():
    with pytest.raises(workbook.WorkbookUnsupported) as excinfo:
        workbook.plan(FakeGraph(), DRIVE, ITEM, SHEET, {"A1": "x"}, "macro.xlsm")
    assert ".xlsx" in excinfo.value.reason


def test_cloning_a_sheet_falls_back():
    with pytest.raises(workbook.WorkbookUnsupported) as excinfo:
        workbook.plan(FakeGraph(), DRIVE, ITEM, "NEW", {"A1": "x"}, "plan.xlsx", copy_sheet_from="SYS2_PIN")
    assert "copy_sheet_from" in excinfo.value.reason


def test_workbook_api_unavailable_falls_back():
    """A Graph 404/403 on the sheet listing means the API cannot serve this file."""
    graph = FakeGraph(fail={"/worksheets?": Mcp365Error("HTTP 403 Forbidden")})
    with pytest.raises(workbook.WorkbookUnsupported) as excinfo:
        workbook.plan(graph, DRIVE, ITEM, SHEET, {"A1": "x"}, "plan.xlsx")
    assert "403" in excinfo.value.reason


def test_a_failed_cell_read_falls_back_rather_than_erroring():
    graph = FakeGraph(fail={"range(address='Q34')": Mcp365Error("HTTP 423 Locked")})
    with pytest.raises(workbook.WorkbookUnsupported):
        workbook.plan(graph, DRIVE, ITEM, SHEET, {"Q34": "x"}, "plan.xlsx")


def test_no_cells_is_an_error_not_a_fallback():
    with pytest.raises(Mcp365Error):
        workbook.plan(FakeGraph(), DRIVE, ITEM, SHEET, {}, "plan.xlsx")


# --------------------------------------------------------------- preview/plan


def test_plan_resolves_the_sheet_by_id_and_only_reads():
    """The id goes in the URL bare and percent-encoded: the quoted form 404s."""
    graph = FakeGraph(values={"Q34": "cũ"})
    sheet_id, changes = workbook.plan(graph, DRIVE, ITEM, SHEET, {"Q34": "mới"}, "plan.xlsx")
    assert sheet_id == SHEET_ID
    assert changes == [{"cell": "Q34", "old": "cũ", "new": "mới"}]
    assert graph.paths() == [
        f"{BASE}/worksheets?$select=id,name",
        f"{BASE}/worksheets/{ENCODED_ID}/range(address='Q34')?$select=values",
    ]
    assert all(c["method"] == "GET" for c in graph.calls), "preview must not write"
    assert not any(c["session_id"] for c in graph.calls), "preview must not open a session"


def test_plan_for_a_missing_sheet_reports_it_will_be_created():
    graph = FakeGraph()
    sheet_id, changes = workbook.plan(graph, DRIVE, ITEM, "Sprint 5", {"A1": "x"}, "plan.xlsx")
    assert sheet_id == ""
    assert changes[0]["cell"] == "(sheet)" and "Sprint 5" in changes[0]["new"]


# ---------------------------------------------------------------------- write


def test_apply_patches_each_cell_inside_one_session():
    graph = FakeGraph()
    written, warnings = workbook.apply(graph, DRIVE, ITEM, SHEET, SHEET_ID, {"Q34": "xong", "R34": 7})
    assert (written, warnings) == (2, [])

    assert graph.paths("POST") == [f"{BASE}/createSession", f"{BASE}/closeSession"]
    patches = [c for c in graph.calls if c["method"] == "PATCH"]
    assert [c["path"] for c in patches] == [
        f"{BASE}/worksheets/{ENCODED_ID}/range(address='Q34')",
        f"{BASE}/worksheets/{ENCODED_ID}/range(address='R34')",
    ]
    assert [c["body"] for c in patches] == [{"values": [["xong"]]}, {"values": [[7]]}]


def test_the_session_id_travels_on_every_call_after_it_opens():
    graph = FakeGraph()
    workbook.apply(graph, DRIVE, ITEM, SHEET, SHEET_ID, {"Q34": "x"})
    after_open = graph.calls[1:]  # everything past createSession
    assert after_open and all(c["session_id"] == "SESSION-1" for c in after_open)
    assert graph.calls[0]["body"] == {"persistChanges": True}


def test_the_session_is_closed_even_when_a_patch_blows_up():
    graph = FakeGraph(fail={"range(address='Q34')": Mcp365Error("HTTP 500")})
    with pytest.raises(Mcp365Error):
        workbook.apply(graph, DRIVE, ITEM, SHEET, SHEET_ID, {"Q34": "x"})
    assert f"{BASE}/closeSession" in graph.paths("POST")


def test_a_refused_session_still_writes_sessionless():
    """The session is an optimisation, not a requirement."""
    graph = FakeGraph(fail={"createSession": Mcp365Error("HTTP 403 Forbidden")})
    written, warnings = workbook.apply(graph, DRIVE, ITEM, SHEET, SHEET_ID, {"Q34": "x"})
    assert written == 1
    assert warnings and "403" in warnings[0]
    assert not any(c["session_id"] for c in graph.calls)


def test_a_failed_close_never_fails_the_write():
    graph = FakeGraph(fail={"closeSession": Mcp365Error("HTTP 500")})
    written, warnings = workbook.apply(graph, DRIVE, ITEM, SHEET, SHEET_ID, {"Q34": "x"})
    assert written == 1 and any("hết hạn" in w for w in warnings)


def test_a_partial_write_says_what_landed_and_is_never_retried_whole_file():
    """Re-uploading the workbook after a cell landed would undo the co-authoring."""
    graph = FakeGraph(fail={"range(address='R34')": Mcp365Error("HTTP 423 Locked")})
    with pytest.raises(Mcp365Error) as excinfo:
        workbook.apply(graph, DRIVE, ITEM, SHEET, SHEET_ID, {"Q34": "a", "R34": "b"})
    assert "1/2" in str(excinfo.value)
    # An Mcp365Error, not a WorkbookUnsupported: the caller must not fall back.
    assert not isinstance(excinfo.value, workbook.WorkbookUnsupported)


def test_a_missing_sheet_is_created_before_the_first_patch():
    graph = FakeGraph()
    written, _ = workbook.apply(graph, DRIVE, ITEM, "Sprint 5", "", {"A1": "x"})
    assert written == 1
    assert graph.paths("POST")[0] == f"{BASE}/worksheets/add"
    assert graph.calls[0]["body"] == {"name": "Sprint 5"}
    assert "/worksheets/%7BNEW-SHEET%7D/range(address='A1')" in graph.paths("PATCH")[0]


def test_a_value_excel_stored_differently_is_reported():
    class Coercing(FakeGraph):
        def __call__(self, path, method="GET", body=None, session_id="", context=""):
            res = super().__call__(path, method, body, session_id, context)
            if method == "PATCH":
                return {**res, "values": [["KHÁC"]]}
            return res

    _written, warnings = workbook.apply(Coercing(), DRIVE, ITEM, SHEET, SHEET_ID, {"Q34": "gửi"})
    assert warnings and "KHÁC" in warnings[0]


def test_render_changes_shows_empty_cells_explicitly():
    out = workbook.render_changes([{"cell": "Q34", "old": "", "new": "xong"}], SHEET)
    assert "`Q34`" in out and "_(trống)_" in out


# ------------------------------------------- which path `update_sharepoint_sheet` takes

import io  # noqa: E402

from mcp.server.mcpserver import MCPServer  # noqa: E402
from mcp.server.mcpserver.exceptions import ToolError  # noqa: E402
from openpyxl import Workbook, load_workbook  # noqa: E402

import tools as tools_mod  # noqa: E402


def _xlsx_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET
    ws["Q34"] = "cũ"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


class FakeSharePoint:
    def __init__(self, name="plan.xlsx", graph=None):
        self.name = name
        self.graph = graph or FakeGraph(values={"Q34": "cũ"})
        self.uploaded: bytes | None = None
        self.downloads = 0

    def resolve_file(self, _ref):
        return DRIVE, {"id": ITEM, "name": self.name, "eTag": "etag-1", "webUrl": "https://sp/plan.xlsx"}

    def call_workbook(self, path, method="GET", body=None, session_id="", context=""):
        return self.graph(path, method, body, session_id, context)

    def read_file_bytes(self, _drive_id, _item):
        self.downloads += 1
        return _xlsx_bytes()

    def put_file_bytes(self, _drive_id, _item_id, data, if_match=""):
        self.uploaded = data
        assert if_match == "etag-1"
        return {"webUrl": "https://sp/plan.xlsx"}


def _server(monkeypatch, fake):
    monkeypatch.setattr(tools_mod, "sp", lambda: fake)
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)
    return mcp


@pytest.mark.anyio
async def test_an_xlsx_is_written_cell_by_cell_and_never_re_uploaded(monkeypatch):
    """The whole point: no download, no whole-file PUT, so an open file is fine."""
    fake = FakeSharePoint()
    mcp = _server(monkeypatch, fake)
    await mcp.call_tool(
        "update_sharepoint_sheet",
        {"file_url_or_guid": "GUID", "sheet": SHEET, "cells": {"Q34": "xong"}, "is_user_confirm": True},
    )
    assert fake.downloads == 0 and fake.uploaded is None
    assert [c["method"] for c in fake.graph.calls if c["method"] == "PATCH"] == ["PATCH"]


@pytest.mark.anyio
async def test_the_preview_call_writes_nothing_and_shows_the_change_list(monkeypatch):
    fake = FakeSharePoint()
    mcp = _server(monkeypatch, fake)
    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "update_sharepoint_sheet",
            {"file_url_or_guid": "GUID", "sheet": SHEET, "cells": {"Q34": "xong"}, "is_user_confirm": False},
        )
    refusal = str(excinfo.value)
    assert "`Q34`" in refusal and "cũ" in refusal and "xong" in refusal
    assert not [c for c in fake.graph.calls if c["method"] != "GET"]
    assert fake.uploaded is None


@pytest.mark.anyio
async def test_cloning_a_sheet_falls_back_to_the_whole_file_path_and_says_why(monkeypatch):
    fake = FakeSharePoint()
    mcp = _server(monkeypatch, fake)
    res = await mcp.call_tool(
        "update_sharepoint_sheet",
        {
            "file_url_or_guid": "GUID",
            "sheet": "Sprint 5",
            "cells": {"A1": "mới"},
            "copy_sheet_from": SHEET,
            "is_user_confirm": True,
        },
    )
    assert fake.downloads == 1 and fake.uploaded is not None
    assert load_workbook(io.BytesIO(fake.uploaded))["Sprint 5"]["A1"].value == "mới"
    assert "copy_sheet_from" in str(res.content)


@pytest.mark.anyio
async def test_the_fallback_preview_warns_it_overwrites_the_whole_file(monkeypatch):
    """A silent downgrade would look like a bug the next time a write hits 423."""
    fake = FakeSharePoint()
    mcp = _server(monkeypatch, fake)
    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "update_sharepoint_sheet",
            {
                "file_url_or_guid": "GUID",
                "sheet": "Sprint 5",
                "cells": {"A1": "mới"},
                "copy_sheet_from": SHEET,
                "is_user_confirm": False,
            },
        )
    refusal = str(excinfo.value)
    assert "ghi đè cả file" in refusal and "copy_sheet_from" in refusal
    assert fake.uploaded is None


@pytest.mark.anyio
async def test_graph_refusing_before_any_write_falls_back_instead_of_failing(monkeypatch):
    fake = FakeSharePoint(graph=FakeGraph(fail={"/worksheets?": Mcp365Error("HTTP 403 Forbidden")}))
    mcp = _server(monkeypatch, fake)
    res = await mcp.call_tool(
        "update_sharepoint_sheet",
        {"file_url_or_guid": "GUID", "sheet": SHEET, "cells": {"Q34": "xong"}, "is_user_confirm": True},
    )
    assert fake.uploaded is not None
    assert "403" in str(res.content)
