"""Chat-list sections (CSA conversationFolders): auth, wire format, 412, drafts and system-section guards."""

from __future__ import annotations

import base64
import copy
import json
import logging
import time
import urllib.parse

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import tools
from common.config import reset_config_cache
from common.errors import AuthExpiredError, ConcurrentEditError, Mcp365Error
from teams import csa

CHAT_A = "19:aaaa@thread.v2"
CHAT_B = "19:bbbb@thread.v2"
STATE = {
    "folderHierarchyVersion": 100,
    "migratedPinnedToFavorites": True,
    "conversationFolderOrder": ["qv", "fav", "u1", "u2", "recent", "u3"],
    "conversationFolders": [
        {"id": "qv", "name": "Quick views", "folderType": "QuickViews", "sortType": "UserDefinedCustomOrder",
         "conversationFolderItems": [{"conversationId": "slice-1", "itemType": "Slice"}]},
        {"id": "fav", "name": "Favorites", "folderType": "Favorites", "sortType": "UserDefinedCustomOrder",
         "conversationFolderItems": [{"conversationId": "48:notes", "threadType": "streamofnotes"}]},
        {"id": "u1", "name": "Dự án", "folderType": "UserCreated", "sortType": "MostRecent",
         "conversationFolderItems": [{"conversationId": CHAT_A, "threadType": "chat"}]},
        {"id": "u2", "name": "Khác", "folderType": "UserCreated", "sortType": "UserDefinedCustomOrder",
         "conversationFolderItems": []},
        {"id": "recent", "name": "Recent", "folderType": "RecentChats", "conversationFolderItems": []},
        {"id": "u3", "name": "Cũ", "folderType": "UserCreated", "conversationFolderItems": []},
        {"id": "gone", "name": "Đã xoá", "folderType": "UserCreated", "isDeleted": True,
         "conversationFolderItems": []},
    ],
}


class FakeAuth:
    def __init__(self):
        self.calls = []

    def get_token(self, force_refresh=False):
        self.calls.append(force_refresh)
        return "csa-token-2" if force_refresh else "csa-token"


@pytest.fixture
def server(monkeypatch):
    """CSA over a recorded fake transport. ``fail`` holds statuses the next writes answer with."""
    rec = {"calls": [], "state": copy.deepcopy(STATE), "fail": [], "get_fail": []}

    def request(url, *, method, headers, data, context, max_retries=None, **_kw):
        body = json.loads(data) if data else None
        rec["calls"].append({"method": method, "url": url, "headers": headers, "body": body,
                             "max_retries": max_retries})
        pending = rec["fail"] if method != "GET" else rec["get_fail"]
        if pending:
            status = pending.pop(0)
            err = (AuthExpiredError if status == 401 else ConcurrentEditError)(f"HTTP {status}")
            err.http_status = status
            raise err
        if method == "GET":
            return 200, json.dumps(rec["state"]).encode(), {}
        return 200, b"{}", {}

    monkeypatch.setattr("teams.csa.request", request)
    rec["client"] = csa.SectionsClient(FakeAuth())
    return rec


def writes(rec):
    return [c for c in rec["calls"] if c["method"] != "GET"]


# --------------------------------------------------------------------- auth


def _jwt(**claims):
    def part(obj):
        return base64.urlsafe_b64encode(json.dumps(obj).encode()).rstrip(b"=").decode()

    return f"{part({'alg': 'none'})}.{part(claims)}.sig"


def test_csa_token_minted_silently_like_mail_and_cached_in_memory_only(monkeypatch, identity, caplog):
    reset_config_cache()
    monkeypatch.setattr("teams.csa.TeamsAuthManager.get_identity", lambda: identity)
    seen = {"mints": 0}
    token = _jwt(aud="https://chatsvcagg.teams.microsoft.com", exp=time.time() + 3600)

    def cookies(domain, use_cache):
        seen["cookie_domain"] = domain
        return {"ESTSAUTHPERSISTENT": "browser-session"}

    def redirect(url, **kwargs):
        seen["authorize"] = url
        seen["redirect"] = kwargs
        return "code-1"

    def token_call(url, **kwargs):
        seen["mints"] += 1
        seen["token_url"], seen["token_request"] = url, kwargs
        return {"access_token": token, "refresh_token": "must-not-be-kept", "expires_in": 3600}

    monkeypatch.setattr("common.spa_auth.ChromeCookieDecryptor.get_cookies_for_domain", cookies)
    monkeypatch.setattr("common.spa_auth.capture_redirect_fragment", redirect)
    monkeypatch.setattr("common.spa_auth.request_json", token_call)
    caplog.set_level(logging.DEBUG)

    auth = csa.CsaAuthManager()
    assert auth.get_token() == token and auth.get_token() == token and seen["mints"] == 1
    authorize = urllib.parse.urlsplit(seen["authorize"])
    params = urllib.parse.parse_qs(authorize.query)
    assert authorize.netloc == "login.microsoftonline.com" and seen["cookie_domain"] == "login.microsoftonline.com"
    assert authorize.path == f"/{identity.tenant_id}/oauth2/v2.0/authorize"
    assert params["client_id"] == ["5e3ce6c0-2b1f-4285-8d4b-75ee78787346"]
    assert params["redirect_uri"] == ["https://teams.cloud.microsoft/v2"]
    assert params["scope"] == ["https://chatsvcagg.teams.microsoft.com/.default openid profile"]
    assert params["prompt"] == ["none"] and params["response_mode"] == ["fragment"]
    assert params["code_challenge_method"] == ["S256"] and params["login_hint"] == [identity.upn]
    assert seen["redirect"]["expected"]["state"] == params["state"][0]
    assert seen["token_request"]["headers"]["Origin"] == "https://teams.cloud.microsoft"
    form = urllib.parse.parse_qs(seen["token_request"]["data"].decode())
    assert form["grant_type"] == ["authorization_code"] and form["code_verifier"][0]
    assert "must-not-be-kept" not in repr(vars(auth))
    assert token not in caplog.text and "browser-session" not in caplog.text


def test_csa_token_config_overrides_and_wrong_audience(monkeypatch, identity):
    monkeypatch.setenv("MCP365_TEAMS_CSA_CLIENT_ID", "other-client")
    monkeypatch.setenv("MCP365_TEAMS_CSA_TENANT_ID", "tenant-x")
    monkeypatch.setenv("MCP365_TEAMS_CSA_REDIRECT_URI", "https://teams.example/v2")
    reset_config_cache()
    monkeypatch.setattr("teams.csa.TeamsAuthManager.get_identity", lambda: identity)
    seen = {}
    monkeypatch.setattr("common.spa_auth.ChromeCookieDecryptor.get_cookies_for_domain",
                        lambda d, use_cache: {"ESTSAUTH": "s"})
    monkeypatch.setattr("common.spa_auth.capture_redirect_fragment", lambda url, **kw: seen.update(url=url) or "c")
    monkeypatch.setattr("common.spa_auth.request_json",
                        lambda url, **kw: {"access_token": _jwt(aud="https://outlook.office.com", exp=time.time() + 60)})
    with pytest.raises(AuthExpiredError, match="không dành cho Teams"):
        csa.CsaAuthManager().get_token()
    params = urllib.parse.parse_qs(urllib.parse.urlsplit(seen["url"]).query)
    assert "/tenant-x/" in seen["url"] and params["client_id"] == ["other-client"]
    assert params["redirect_uri"] == ["https://teams.example/v2"]
    monkeypatch.setattr("common.spa_auth.ChromeCookieDecryptor.get_cookies_for_domain", lambda d, use_cache: {})
    with pytest.raises(AuthExpiredError, match="Stay signed in"):
        csa.CsaAuthManager().get_token()


# --------------------------------------------------------------------- wire


def test_get_url_query_bearer_and_401_refresh(server):
    state = server["client"].state()
    assert state["folderHierarchyVersion"] == 100
    call = server["calls"][0]
    parts = urllib.parse.urlsplit(call["url"])
    assert (parts.netloc, parts.path) == ("teams.cloud.microsoft", "/api/csa/api/v1/teams/users/me/conversationFolders")
    assert parts.query == "supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true"
    assert call["headers"]["Authorization"] == "Bearer csa-token" and "Content-Type" not in call["headers"]
    server["get_fail"].append(401)
    server["calls"].clear()
    server["client"].state()
    assert [c["headers"]["Authorization"] for c in server["calls"]] == ["Bearer csa-token", "Bearer csa-token-2"]
    server["get_fail"][:] = [401, 401]
    with pytest.raises(AuthExpiredError) as exc:
        server["client"].state()
    assert "teams.cloud.microsoft" in exc.value.remediation


def test_write_carries_latest_version_and_actions_without_retrying_5xx(server):
    client = server["client"]
    state = client.state()
    write = csa.Write("POST", body={"actions": [{"action": "AddItem", "folderId": "u2", "itemId": CHAT_B}]}, safe=True)
    client.apply(state, lambda s: write, "x")
    (call,) = writes(server)
    assert call["body"] == {"folderHierarchyVersion": 100, "supportsAdditionalSystemGeneratedFolders": True,
                            "supportsSliceItems": True, "actions": write.body["actions"]}
    assert call["headers"]["Content-Type"] == "application/json" and call["max_retries"] == 0
    assert "?supportsAdditionalSystemGeneratedFolders=true&supportsSliceItems=true" in call["url"]
    rename = csa.Write("PUT", "u2", {"folderName": "Mới"}, safe=True)
    client.apply(state, lambda s: rename, "x")
    assert writes(server)[-1]["url"].split("?")[0].endswith("/conversationFolders/u2")


def test_412_retries_safe_write_once_with_fresh_version(server):
    client = server["client"]
    state = client.state()
    server["state"]["folderHierarchyVersion"] = 101
    server["fail"].append(412)

    def plan(s):
        actions, _ = csa.move_actions(s, [CHAT_B], csa.find_section(s, "u2"))
        return csa.Write("POST", body={"actions": actions}, safe=True)

    client.apply(state, plan, "x")
    assert [c["body"]["folderHierarchyVersion"] for c in writes(server)] == [100, 101]
    # A second 412 is a conflict, not a loop.
    server["calls"].clear()
    server["fail"][:] = [412, 412]
    with pytest.raises(ConcurrentEditError, match="412"):
        client.apply(state, plan, "x")
    assert len(writes(server)) == 2


def test_412_on_safe_write_whose_plan_changed_sends_nothing_more(server):
    client = server["client"]
    state = client.state()
    server["state"]["conversationFolders"][3]["conversationFolderItems"].append({"conversationId": CHAT_B})
    server["fail"].append(412)

    def plan(s):
        actions, _ = csa.move_actions(s, [CHAT_B], csa.find_section(s, "u2"))
        return csa.Write("POST", body={"actions": actions}, safe=True)

    with pytest.raises(ConcurrentEditError, match="khác bản nháp"):
        client.apply(state, plan, "x")
    assert len(writes(server)) == 1


@pytest.mark.parametrize("write", [
    csa.Write("POST", body={"actions": [{"action": "ReorderFolders", "newFoldersOrder": ["u2", "u1"]}]}),
    csa.Write("DELETE", "u2"),
])
def test_412_on_reorder_or_delete_is_a_conflict(server, write):
    client = server["client"]
    state = client.state()
    server["fail"].append(412)
    with pytest.raises(ConcurrentEditError, match="không tự thử lại"):
        client.apply(state, lambda s: write, "x")
    assert len(writes(server)) == 1 and len(server["calls"]) == 2


def test_moves_name_the_section_a_chat_leaves():
    state = copy.deepcopy(STATE)
    actions, lines = csa.move_actions(state, [CHAT_A, CHAT_B, "48:notes"], csa.find_section(state, "Khác"))
    assert actions == [
        {"action": "AddItem", "folderId": "u2", "itemId": CHAT_A}, {"action": "RemoveItem", "folderId": "u1", "itemId": CHAT_A},
        {"action": "AddItem", "folderId": "u2", "itemId": CHAT_B},
        {"action": "AddItem", "folderId": "u2", "itemId": "48:notes"}, {"action": "RemoveItem", "folderId": "fav", "itemId": "48:notes"},
    ]
    assert "**chuyển** khỏi mục `Dự án`" in lines[0] and "chuyển" not in lines[1] and "`Favorites`" in lines[2]
    actions, lines = csa.move_actions(state, [CHAT_A, CHAT_B], None, remove=True)
    assert actions == [{"action": "RemoveItem", "folderId": "u1", "itemId": CHAT_A}] and "bỏ qua" in lines[1]
    assert csa.move_actions(state, [CHAT_A], csa.find_section(state, "u1"))[0] == []


def test_section_lookup_is_exact_and_guards_system_sections():
    state = copy.deepcopy(STATE)
    assert csa.find_section(state, "du an")["id"] == "u1"
    assert csa.find_section(state, "favorites")["id"] == "fav"
    for query in ("Dự", "Đã xoá"):
        with pytest.raises(Mcp365Error, match="Không có mục"):
            csa.find_section(state, query)
    with pytest.raises(Mcp365Error, match="mục hệ thống"):
        csa.find_section(state, "Quick views")
    with pytest.raises(Mcp365Error, match="mục hệ thống"):
        csa.find_section(state, "fav", ("UserCreated",))


# --------------------------------------------------------------------- tools


class FakeTeams:
    def __init__(self, identity):
        self.identity = identity
        self.listings = 0

    def find_conversation(self, ident):
        return {"48:notes": {"id": "48:notes", "name": "Chat with yourself (Notes)"},
                "A": {"id": CHAT_A, "name": "Nhóm A"}, "B": {"id": CHAT_B, "name": "Nhóm B"}}[ident]

    def list_conversations(self, page_size=50):
        self.listings += 1
        return [{"id": CHAT_A, "name": "Nhóm A"}]


@pytest.fixture
def mcp(server, monkeypatch, identity):
    fake = FakeTeams(identity)
    monkeypatch.setattr(tools, "sections", lambda: server["client"])
    monkeypatch.setattr(tools, "teams", lambda: fake)
    app = MCPServer("t")
    tools.register_teams_tools(app)
    server["teams"] = fake
    return app


async def _draft(mcp, name, args):
    with pytest.raises(ToolError) as exc:
        await mcp.call_tool(name, {**args, "is_user_confirm": False})
    return str(exc.value)


@pytest.mark.anyio
async def test_list_hides_system_sections_except_favorites_and_names_chats_from_one_listing(mcp, server):
    text = str(await mcp.call_tool("list_teams_sections", {}))
    assert text.index("Favorites") < text.index("Dự án") < text.index("Khác") < text.index("Cũ")
    assert "Quick views" not in text and "**Recent**" not in text and "Đã xoá" not in text
    assert "Nhóm A (`19:aaaa@thread.v2`)" in text and "Chat with yourself (Notes) (`48:notes`)" in text
    assert "Ẩn 2 mục hệ thống (QuickViews, RecentChats)" in text
    assert server["teams"].listings == 1 and writes(server) == []
    text = str(await mcp.call_tool("list_teams_sections", {"include_system": True}))
    assert text.index("Quick views") < text.index("Favorites") and "slice-1" in text and "Slice" in text


@pytest.mark.anyio
@pytest.mark.parametrize("name,args,expected", [
    ("manage_teams_section", {"action": "create", "name": "zz-test"}, "Mục mới `zz-test`"),
    ("manage_teams_section", {"action": "rename", "section_name_or_id": "u2", "name": "Mới"}, "`Khác` → `Mới`"),
    ("manage_teams_section", {"action": "delete", "section_name_or_id": "Dự án"},
     "KHÔNG chat nào bị xoá"),
    ("manage_teams_section", {"action": "reorder", "section_order": ["u3", "Dự án", "Khác"]},
     "1. Quick views (giữ chỗ)\n2. Favorites (giữ chỗ)\n3. Cũ\n4. Dự án\n5. Recent (giữ chỗ)\n6. Khác"),
    ("set_teams_chat_section", {"chat_name_or_id": "A", "section_name_or_id": "Khác"}, "**chuyển** khỏi mục `Dự án`"),
    ("set_teams_chat_section", {"chat_names_or_ids": ["A", "48:notes"], "remove": True}, "bỏ khỏi mục `Favorites`"),
    ("set_teams_chat_preferences", {"chat_name_or_id": "48:notes", "pin": False}, "pin=False (bỏ ghim"),
])
async def test_drafts_send_no_writes(mcp, server, name, args, expected):
    text = await _draft(mcp, name, args)
    assert "CHƯA GỬI" in text and expected in text
    assert writes(server) == [] and all(c["method"] == "GET" for c in server["calls"])


@pytest.mark.anyio
async def test_delete_draft_lists_chats_that_return_to_the_normal_list(mcp, server):
    text = await _draft(mcp, "manage_teams_section", {"action": "delete", "section_name_or_id": "u1"})
    assert "quay về danh sách thường" in text and "Nhóm A (`19:aaaa@thread.v2`)" in text


@pytest.mark.anyio
@pytest.mark.parametrize("args", [
    {"action": "rename", "section_name_or_id": "Favorites", "name": "x"},
    {"action": "delete", "section_name_or_id": "qv"},
    {"action": "reorder", "section_order": ["fav", "u1", "u2", "u3"]},
    {"action": "reorder", "section_order": ["u1", "u2"]},
    {"action": "create", "name": "dự án"},
])
async def test_system_sections_and_bad_plans_are_refused_before_any_write(mcp, server, args):
    with pytest.raises(ToolError) as exc:
        await mcp.call_tool("manage_teams_section", {**args, "is_user_confirm": True})
    assert "CHƯA GỬI" not in str(exc.value) and writes(server) == []


@pytest.mark.anyio
async def test_approved_section_writes_match_the_draft(mcp, server):
    args = {"action": "reorder", "section_order": ["u3", "u1", "u2"]}
    with pytest.raises(ToolError, match="Thiếu bản nháp"):
        await mcp.call_tool("manage_teams_section", {**args, "is_user_confirm": True})
    await _draft(mcp, "manage_teams_section", args)
    assert "Đã sắp lại" in str(await mcp.call_tool("manage_teams_section", {**args, "is_user_confirm": True}))
    (call,) = writes(server)
    assert call["body"]["actions"] == [{"action": "ReorderFolders",
                                        "newFoldersOrder": ["qv", "fav", "u3", "u1", "recent", "u2"]}]
    # The hierarchy changed between draft and approval: nothing is written.
    server["calls"].clear()
    args = {"chat_name_or_id": "B", "section_name_or_id": "u2"}
    await _draft(mcp, "set_teams_chat_section", args)
    server["state"]["conversationFolders"][1]["conversationFolderItems"].append({"conversationId": CHAT_B})
    with pytest.raises(ToolError, match="đã thay đổi"):
        await mcp.call_tool("set_teams_chat_section", {**args, "is_user_confirm": True})
    assert writes(server) == []


@pytest.mark.anyio
async def test_create_and_pin_bodies(mcp, server):
    args = {"action": "create", "name": "zz-test"}
    await _draft(mcp, "manage_teams_section", args)
    await mcp.call_tool("manage_teams_section", {**args, "is_user_confirm": True})
    assert writes(server)[-1]["body"]["actions"] == [{"action": "AddFolder"}]
    assert writes(server)[-1]["body"]["folderName"] == "zz-test"
    args = {"chat_name_or_id": "A", "pin": True}
    text = await _draft(mcp, "set_teams_chat_preferences", args)
    assert "**chuyển** khỏi mục `Dự án`" in text
    assert "đã ghim" in str(await mcp.call_tool("set_teams_chat_preferences", {**args, "is_user_confirm": True}))
    assert writes(server)[-1]["body"]["actions"] == [
        {"action": "AddItem", "folderId": "fav", "itemId": CHAT_A}, {"action": "RemoveItem", "folderId": "u1", "itemId": CHAT_A}]
    with pytest.raises(ToolError, match="Không có thay đổi"):
        await mcp.call_tool("set_teams_chat_preferences", {"chat_name_or_id": "48:notes", "pin": True,
                                                           "is_user_confirm": False})


def test_health_reports_csa_channel(monkeypatch, server):
    from common import health

    monkeypatch.setattr("teams.csa.CsaAuthManager.get_token", lambda self, force_refresh=False: "csa-token")
    probe = health._probe_csa()
    assert probe["status"] == "✅" and "6 mục (3 mục bạn tạo)" in probe["detail"]

    def refuse(self, force_refresh=False):
        raise AuthExpiredError("Chrome không có phiên đăng nhập Microsoft bền để mở Teams (mục chat).", "x")

    monkeypatch.setattr("teams.csa.CsaAuthManager.get_token", refuse)
    probe = health._probe_csa()
    assert probe["status"] == "⚠️" and "teams.cloud.microsoft" in probe["fix"] and "Stay signed in" in probe["fix"]
