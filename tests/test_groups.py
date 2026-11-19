"""Roster writes are bound to resolved identities; partial service responses are not success."""

import io
import json
import urllib.error
from urllib.parse import quote

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import tools
from common.errors import Mcp365Error, classify_http_error
from teams import groups
from teams.client import TeamsClient

CHAT = "19:example@thread.v2"
A = "8:orgid:11111111-1111-4111-8111-111111111111"
B = "8:orgid:22222222-2222-4222-8222-222222222222"


@pytest.fixture
def wire(monkeypatch, identity):
    c = TeamsClient()
    monkeypatch.setattr(type(c), "identity", property(lambda self: identity))
    monkeypatch.setattr(c, "_auth", lambda: {"region": "apac", "token": "test"})
    calls = []
    response = [200, b"{}", {}]

    def request(url, *, headers, method, data, **kwargs):
        assert headers["Authentication"] == "skypetoken=test"
        if data is not None:
            assert headers["Content-Type"] == "application/json"
        calls.append((method, url.split("/v1", 1)[1], json.loads(data) if data else None, headers))
        return tuple(response)

    monkeypatch.setattr("teams.client.request", request)
    return c, calls, response


def test_create_id_from_location_no_first_message(wire):
    c, calls, response = wire
    c._conv_cache = []
    response[:] = 201, b"", {"Location": "https://service/threads/" + quote(CHAT, safe="")}
    assert groups.create(c, [A], "topic") == CHAT
    method, url, body, headers = calls.pop()
    assert not calls and (method, url, headers["Prefer"]) == ("POST", "/threads", "return=representation")
    assert body == {"properties": {"threadType": "chat", "chatFilesIndexId": "2", "topic": "topic"}, "members": [{"id": c.identity.mri, "role": "Admin"}, {"id": A, "role": "Admin"}]}
    assert c._conv_cache is None


@pytest.mark.parametrize("days,expected", [(None, -1), (0, 0), (2, 827200000)])
def test_add_history_and_partial_response(wire, monkeypatch, days, expected):
    """Like Teams Web: no history = an explicit -1 (never omitted), all = 0, N days = now - N days in ms."""
    c, calls, response = wire
    monkeypatch.setattr("teams.groups.time.time", lambda: 1000000)
    response[:] = 207, json.dumps({"multipleStatus": {A: 200, B: {"statusCode": 403}}}).encode(), {}
    report = groups.add(c, CHAT, [A, B], days)
    method, url, body, _ = calls.pop()
    assert (method, url) == ("PUT", groups.group_path(CHAT) + "/members")
    assert body["members"] == [{"id": m, "role": "Admin", "shareHistoryTime": expected} for m in (A, B)]
    assert f"{A}: đã thêm (200)" in report and f"{B}: bị từ chối (403)" in report


def test_negative_history_days_is_refused_before_any_request(wire):
    c, calls, _ = wire
    with pytest.raises(Mcp365Error, match="không được âm"):
        groups.add(c, CHAT, [A], -1)
    assert not calls


def test_add_403_with_member_details(wire, monkeypatch):
    c, _, _ = wire
    raw = json.dumps({"error": {"multipleStatus": [{"id": A, "code": 403}]}}).encode()
    error = classify_http_error(urllib.error.HTTPError("https://test", 403, "Forbidden", {}, io.BytesIO(raw)))
    error.http_status = 403
    monkeypatch.setattr(c, "_chat", lambda *a, **kw: (_ for _ in ()).throw(error))
    with pytest.raises(Mcp365Error, match=f"{A}: bị từ chối"):
        groups.add(c, CHAT, [A], None)


def test_remove_partial_and_self_guard(wire, monkeypatch):
    c, _, _ = wire
    calls = []

    def remove(method, url, **kw):
        calls.append((method, url))
        if url.endswith(quote(B, safe="")):
            raise Mcp365Error("HTTP 403")
        return 204, b"", {}

    monkeypatch.setattr(c, "_chat", remove)
    with pytest.raises(Mcp365Error, match="leave_teams_chat"):
        groups.remove(c, CHAT, [A, c.identity.mri])
    assert not calls
    c._conv_cache = []
    report = groups.remove(c, CHAT, [A, B])
    assert sorted(calls) == [("DELETE", groups.group_path(CHAT) + "/members/" + quote(m, safe="")) for m in (A, B)]
    assert f"{A}: đã xoá khỏi nhóm" in report and f"{B}: lỗi: HTTP 403" in report and c._conv_cache is None
    assert "đã rời nhóm" not in report
    calls.clear()
    assert groups.remove(c, CHAT, [c.identity.mri], leave=True) == f"{c.identity.mri}: đã rời nhóm"


def test_rename_leave_and_personal_preferences(wire):
    c, calls, _ = wire
    groups.rename(c, CHAT, "")
    assert calls.pop()[:3] == ("PUT", groups.group_path(CHAT) + "/properties?name=topic", {"topic": ""})
    groups.remove(c, CHAT, [c.identity.mri], leave=True)
    assert calls.pop()[:3] == ("DELETE", groups.group_path(CHAT) + "/members/" + quote(c.identity.mri, safe=""), None)
    groups.preferences(c, "48:notes", True, False)
    prefix = "/users/ME/conversations/48%3Anotes/properties?name="
    assert [c[:3] for c in calls] == [("PUT", prefix + k, {k: v}) for k, v in [("alerts", "false"), ("unpinnedTime", None), ("historyHiddenTime", None)]]


@pytest.mark.anyio
@pytest.mark.parametrize("name,args", [
    ("create_teams_group_chat", {"members": [A], "topic": "New"}),
    ("add_teams_group_members", {"chat_name_or_id": CHAT, "user": A, "share_history_days": 0}),
    ("remove_teams_group_members", {"chat_name_or_id": CHAT, "user": A}),
    ("update_teams_group_chat", {"chat_name_or_id": CHAT, "topic": "New"}),
    ("leave_teams_chat", {"chat_name_or_id": CHAT}),
    ("set_teams_chat_preferences", {"chat_name_or_id": "48:notes", "hide": True}),
])
async def test_drafts_never_issue_mutations(wire, monkeypatch, name, args):
    c, calls, _ = wire
    monkeypatch.setattr(tools, "teams", lambda: c)
    monkeypatch.setattr(c, "resolve_member", lambda _: {"mri": A, "name": "Person A"})
    server = MCPServer("test")
    tools.register_teams_tools(server)
    with pytest.raises(ToolError, match="CHƯA GỬI") as exc:
        await server.call_tool(name, {**args, "is_user_confirm": False})
    assert all(call[0] == "GET" for call in calls)
    if "user" in args or "members" in args:
        assert A in str(exc.value) and "Person A" in str(exc.value)


@pytest.mark.anyio
async def test_changed_resolution_refuses_approved_write(wire, monkeypatch):
    c, calls, _ = wire
    monkeypatch.setattr(tools, "teams", lambda: c)
    choice = iter([A, B])
    monkeypatch.setattr(c, "resolve_member", lambda _: {"mri": next(choice), "name": "Same Name"})
    server = MCPServer("test")
    tools.register_teams_tools(server)
    args = {"user": "Same Name", "chat_name_or_id": CHAT, "share_history_days": 0}
    with pytest.raises(ToolError, match="toàn bộ lịch sử"):
        await server.call_tool("add_teams_group_members", {**args, "is_user_confirm": False})
    with pytest.raises(ToolError, match="đã thay đổi"):
        await server.call_tool("add_teams_group_members", {**args, "is_user_confirm": True})
    assert all(call[0] == "GET" for call in calls)


def test_namesakes_and_unknown_member_are_refused(wire, monkeypatch):
    c, _, _ = wire
    monkeypatch.setattr(c, "search_users", lambda query, max_results: [{"teams_mri": m, "name": "Same Name"} for m in (A, B)])
    with pytest.raises(Mcp365Error) as exc:
        c.resolve_member("Same Name")
    assert A in str(exc.value) and B in str(exc.value)
    with pytest.raises(Mcp365Error, match="Không xác định"):
        c.resolve_member("No Such Person")
