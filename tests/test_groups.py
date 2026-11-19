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


# ------------------------------------------ leave meeting chats, delete 1:1 chats (19/11)

MEETING = "19:meeting_abc@thread.v2"
ME_GUID = "b6cf511d-9f31-4a84-89d8-3a400a1a544f"
PEER_GUID = "a2300b55-0000-4000-8000-000000000001"
BOT_GUID = "9a961b11-0000-4000-8000-000000000002"
DM = f"19:{PEER_GUID}_{ME_GUID}@unq.gbl.spaces"
BOT_DM = f"19:{ME_GUID}_{BOT_GUID}@unq.gbl.spaces"  # Teams keeps the user first in a bot chat


class FakeSections:
    """CSA stand-in: one custom section holding ``held`` chats; ``apply_each`` records the requests."""

    def __init__(self, held=()):
        self.st = {"folderHierarchyVersion": 1, "migratedPinnedToFavorites": True, "conversationFolderOrder": ["u1"],
                   "conversationFolders": [{"id": "u1", "name": "Việc", "folderType": "UserCreated",
                                            "conversationFolderItems": [{"conversationId": c} for c in held]}]}
        self.sent = []

    def state(self):
        return json.loads(json.dumps(self.st))

    def apply_each(self, state, keys, steps, context, sleep=None):
        for key in keys:
            self.sent += steps(state, key)
        return [(k, "") for k in keys]


@pytest.fixture
def chat(monkeypatch, identity):
    """Chat Service routed by path: ``routes[(method, path-prefix)] = (status, body)``; calls recorded."""
    c = TeamsClient()
    monkeypatch.setattr(type(c), "identity", property(lambda self: identity))
    monkeypatch.setattr(c, "_auth", lambda: {"region": "apac", "token": "test", "middle_tier_token": "mt"})
    rec = {"calls": [], "routes": {}}

    def request(url, *, headers, method, data, **kwargs):
        path = url.split("/v1", 1)[1] if "/v1/" in url else url
        rec["calls"].append((method, path, json.loads(data) if data else None))
        for (m, prefix), (status, body) in rec["routes"].items():
            if m == method and path.startswith(prefix):
                if status >= 400:
                    err = Mcp365Error(f"HTTP {status}")
                    err.http_status = status
                    raise err
                return status, json.dumps(body).encode(), {}
        return 200, b"{}", {}

    monkeypatch.setattr("teams.client.request", request)
    monkeypatch.setattr("teams.client.request_json", lambda url, **kw: rec["profiles"](url, **kw))
    rec["profiles"] = lambda url, **kw: []
    rec["client"] = c
    return rec


def _server(chat, monkeypatch, sections):
    monkeypatch.setattr(tools, "teams", lambda: chat["client"])
    monkeypatch.setattr(tools, "sections", lambda: sections)
    app = MCPServer("test")
    tools.register_teams_tools(app)
    return app


def mutations(chat):
    return [c for c in chat["calls"] if c[0] != "GET"]


def test_meeting_chats_pass_only_for_leaving():
    with pytest.raises(Mcp365Error, match="nhóm chat thường"):
        groups.group_path(MEETING)
    assert groups.group_path(MEETING, meeting=True) == "/threads/" + quote(MEETING, safe="")
    with pytest.raises(Mcp365Error):
        groups.group_path(DM, meeting=True)


@pytest.mark.anyio
@pytest.mark.parametrize("name,args", [
    ("add_teams_group_members", {"user": A}),
    ("remove_teams_group_members", {"user": A}),
    ("update_teams_group_chat", {"topic": "x"}),
])
async def test_roster_and_rename_still_refuse_meeting_chats(chat, monkeypatch, name, args):
    app = _server(chat, monkeypatch, FakeSections())
    monkeypatch.setattr(chat["client"], "resolve_member", lambda _: {"mri": A, "name": "Person A"})
    with pytest.raises(ToolError, match="nhóm chat thường") as exc:
        await app.call_tool(name, {"chat_name_or_id": MEETING, **args, "is_user_confirm": False})
    assert "CHƯA GỬI" not in str(exc.value) and chat["calls"] == []


@pytest.mark.anyio
async def test_leave_meeting_chat_then_hide_and_take_out_of_sections(chat, monkeypatch, identity):
    sections = FakeSections(held=[MEETING])
    app = _server(chat, monkeypatch, sections)
    thread = "/threads/" + quote(MEETING, safe="")
    chat["routes"][("GET", thread)] = (200, {"properties": {"threadType": "meeting", "topic": "Họp", "creator": A}})
    args = {"chat_name_or_id": MEETING, "hide": True}
    with pytest.raises(ToolError) as exc:
        await app.call_tool("leave_teams_chat", {**args, "is_user_confirm": False})
    draft = str(exc.value)
    assert "rời chat cuộc họp" in draft and "KHÔNG tự vào lại được" in draft and "mọi thành viên thấy" in draft
    assert "ẩn chat và gỡ khỏi mục `Việc`" in draft and mutations(chat) == []
    text = str(await app.call_tool("leave_teams_chat", {**args, "is_user_confirm": True}))
    assert mutations(chat)[0][:2] == ("DELETE", thread + "/members/" + quote(identity.mri, safe=""))
    assert [m[1].rsplit("=", 1)[1] for m in mutations(chat)[1:]] == ["unpinnedTime", "historyHiddenTime"]
    assert sections.sent == [[{"action": "RemoveItem", "folderId": "u1", "itemId": MEETING}]]
    assert "Đã rời chat" in text and "đã gỡ khỏi mục `Việc`" in text


@pytest.mark.anyio
async def test_meeting_organizer_cannot_leave_like_in_teams(chat, monkeypatch, identity):
    app = _server(chat, monkeypatch, FakeSections())
    chat["routes"][("GET", "/threads/")] = (200, {"properties": {"threadType": "meeting", "creator": identity.mri}})
    with pytest.raises(ToolError, match="người tạo cuộc họp"):
        await app.call_tool("leave_teams_chat", {"chat_name_or_id": MEETING, "is_user_confirm": False})
    assert mutations(chat) == []


def _dm_routes(chat, conv, members, messages=()):
    q = quote(conv["id"], safe="")
    chat["routes"][("GET", f"/users/ME/conversations/{q}/messages")] = (200, {"messages": list(messages)})
    chat["routes"][("GET", f"/users/ME/conversations/{q}")] = (200, conv)
    chat["routes"][("GET", f"/threads/{q}")] = (200, {"members": [{"id": m} for m in members]})


def _msg(sender, name, when="2026-07-27T08:01:36.7310000Z"):
    return {"from": f"https://x/v1/users/ME/contacts/{sender}", "imdisplayname": name, "messagetype": "RichText/Html",
            "originalarrivaltime": when}


@pytest.mark.anyio
async def test_delete_chat_draft_then_sections_clear_history_and_hide(chat, monkeypatch, identity):
    sections = FakeSections(held=[DM])
    app = _server(chat, monkeypatch, sections)
    peer = f"8:orgid:{PEER_GUID}"
    conv = {"id": DM, "lastMessage": _msg(peer, "Người A"), "properties": {}}
    _dm_routes(chat, conv, [identity.mri, peer])
    chat["profiles"] = lambda url, **kw: [{"mri": peer, "displayName": "Người A", "department": "Phòng X"}]
    with pytest.raises(ToolError) as exc:
        await app.call_tool("delete_teams_chat", {"chat_name_or_id": DM, "is_user_confirm": False})
    draft = str(exc.value)
    assert "Bên kia: người: Người A, Phòng X" in draft and "Gỡ khỏi mục `Việc`" in draft
    assert "KHÔNG hoàn tác được" in draft and "19/11" in draft and "bên kia vẫn giữ toàn bộ chat" in draft
    assert mutations(chat) == [] and sections.sent == []
    text = str(await app.call_tool("delete_teams_chat", {"chat_name_or_id": DM, "is_user_confirm": True}))
    q = quote(DM, safe="")
    clear, *hide = mutations(chat)
    # Teams Web: last message time + 1 ms, as an integer.
    assert clear == ("PUT", f"/users/ME/conversations/{q}/properties?name=clearHistoryTime",
                     {"clearHistoryTime": 1785139296732})
    assert [h[1].rsplit("=", 1)[1] for h in hide] == ["unpinnedTime", "historyHiddenTime"]
    assert sections.sent == [[{"action": "RemoveItem", "folderId": "u1", "itemId": DM}]]
    assert "Đã xoá chat phía bạn" in text


@pytest.mark.anyio
async def test_delete_chat_names_a_bot_and_an_unknown_party(chat, monkeypatch, identity):
    app = _server(chat, monkeypatch, FakeSections())
    bot = f"28:{BOT_GUID}"
    conv = {"id": BOT_DM, "lastMessage": _msg(identity.mri, "Tôi"), "properties": {"clearHistoryTime": "1795092989494"}}
    # The id is not in sorted order: only the chat list proves it is the real one (a bot chat).
    chat["routes"][("GET", "/users/ME/conversations?")] = (200, {"conversations": [conv]})
    _dm_routes(chat, conv, [identity.mri, bot], [_msg(identity.mri, "Tôi"), _msg(bot, "V-IT")])
    with pytest.raises(ToolError) as exc:
        await app.call_tool("delete_teams_chat", {"chat_name_or_id": BOT_DM, "is_user_confirm": False})
    assert f"Bên kia: bot / ứng dụng: V-IT (`{bot}`)" in str(exc.value)
    assert "đã bị xoá phía bạn trước đó" in str(exc.value)
    await app.call_tool("delete_teams_chat", {"chat_name_or_id": BOT_DM, "is_user_confirm": True})
    # Deleting again never moves clearHistoryTime back.
    assert mutations(chat)[0][2] == {"clearHistoryTime": 1795092989494}
    # Locked account: gone from the roster, no message of theirs in sight.
    chat["calls"].clear()
    conv = {"id": DM, "lastMessage": _msg(identity.mri, "Tôi"), "properties": {}}
    _dm_routes(chat, conv, [identity.mri])
    with pytest.raises(ToolError, match="không xác định được là người hay bot"):
        await app.call_tool("delete_teams_chat", {"chat_name_or_id": DM, "is_user_confirm": False})


@pytest.mark.anyio
@pytest.mark.parametrize("target,error", [
    ("48:notes", "48:notes"), ("notes", "48:notes"), (CHAT, "chỉ xoá chat 1:1"), (MEETING, "chỉ xoá chat 1:1"),
])
async def test_delete_chat_refuses_notes_and_non_direct_chats(chat, monkeypatch, target, error):
    app = _server(chat, monkeypatch, FakeSections())
    with pytest.raises(ToolError, match=error) as exc:
        await app.call_tool("delete_teams_chat", {"chat_name_or_id": target, "is_user_confirm": True})
    assert "CHƯA GỬI" not in str(exc.value) and chat["calls"] == []


def test_clear_history_refuses_notes_in_code(chat):
    with pytest.raises(Mcp365Error, match="48:notes"):
        groups.clear_history(chat["client"], "48:notes", 1)
    with pytest.raises(Mcp365Error, match="48:notes"):
        groups.clear_history(chat["client"], "48:notifications", 1)
    assert chat["calls"] == []


def test_delete_stamp_follows_teams_web():
    assert groups.delete_stamp({"lastMessage": {"originalarrivaltime": "2026-07-27T08:01:36.731Z"}}, 5) == 1785139296732
    assert groups.delete_stamp({}, 5) == 5
    assert groups.delete_stamp({"properties": {"clearHistoryTime": "9"}}, 5) == 9
