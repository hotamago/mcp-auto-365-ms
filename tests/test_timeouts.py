"""Per-call timeouts: kind defaults, tool override, ceiling, streaming downloads."""

from __future__ import annotations

import http.client

import pytest
from mcp.server.mcpserver import MCPServer

import tools as tools_mod
from common import http as http_mod
from common.config import reset_config_cache
from common.errors import Mcp365Error, TransportError
from common.http import kind_scope, request_to_file, resolve_timeout, timeout_scope


class _Response:
    def __init__(self, body=b"{}", chunks=None, fail_after=None):
        self.status, self.headers = 200, {}
        self._chunks = list(chunks) if chunks is not None else [body]
        self._fail_after = fail_after
        self._reads = 0

    def read(self, n=None):
        if n is None:
            return b"".join(self._chunks)
        self._reads += 1
        if self._fail_after is not None and self._reads > self._fail_after:
            raise http.client.IncompleteRead(b"")
        return self._chunks.pop(0) if self._chunks else b""

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


@pytest.fixture
def seen(monkeypatch):
    """Record the timeout of every urlopen call."""
    timeouts: list[float] = []

    def fake(req, timeout=None):
        timeouts.append(timeout)
        return _Response()

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    monkeypatch.setattr(http_mod.time, "sleep", lambda _s: None)
    return timeouts


# ------------------------------------------------------------ resolution


def test_chat_and_transfer_defaults_differ():
    assert resolve_timeout("chat") == 10.0
    assert resolve_timeout("transfer") == 120.0
    assert resolve_timeout() == 30.0


def test_tool_timeout_overrides_every_kind_and_is_capped():
    with timeout_scope(300):
        assert resolve_timeout("chat") == 300.0
        assert resolve_timeout("transfer") == 300.0
    with timeout_scope(5000):
        assert resolve_timeout("transfer") == 600.0  # never beyond http.timeout_max
    with timeout_scope(None):
        assert resolve_timeout("chat") == 10.0
    assert resolve_timeout("chat") == 10.0  # scope does not leak


def test_explicit_timeout_wins_over_the_tool_override():
    """Latency probes must stay short even when a tool asked for minutes."""
    with timeout_scope(300):
        assert resolve_timeout("chat", explicit=4.0) == 4.0


def test_timeout_defaults_and_ceiling_come_from_env(monkeypatch):
    monkeypatch.setenv("MCP365_HTTP_TIMEOUT_TRANSFER", "200")
    monkeypatch.setenv("MCP365_HTTP_TIMEOUT_MAX", "250")
    reset_config_cache()
    assert resolve_timeout("transfer") == 200.0
    with timeout_scope(1000):
        assert resolve_timeout("chat") == 250.0


def test_requests_pick_up_kind_and_override(seen):
    http_mod.request_bytes("https://example.invalid/a")
    with kind_scope("transfer"):
        http_mod.request_bytes("https://example.invalid/b")
    with timeout_scope(77):
        http_mod.request_bytes("https://example.invalid/c")
    assert seen == [30.0, 120.0, 77.0]


# ------------------------------------------------------------ messages


def test_timeout_advice_depends_on_what_timed_out(monkeypatch):
    def fake(req, timeout=None):
        raise TimeoutError("timed out")

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(TransportError) as transfer:
        http_mod.request("https://example.invalid", kind="transfer", max_retries=0)
    assert "timeout_seconds" in transfer.value.remediation

    with pytest.raises(TransportError) as chat:
        http_mod.request("https://example.invalid", kind="chat", method="POST", max_retries=0)
    assert "timeout_seconds" not in chat.value.remediation
    assert "kiểm tra đã gửi được chưa" in chat.value.remediation


# ------------------------------------------------------------ streaming


def test_download_is_streamed_to_disk(monkeypatch, tmp_path):
    seen_timeout = []

    def fake(req, timeout=None):
        seen_timeout.append(timeout)
        return _Response(chunks=[b"ab", b"cd", b"e"])

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    dest = tmp_path / "rec.mp4"
    assert request_to_file("https://example.invalid/rec.mp4", dest, chunk_size=2) == 5
    assert dest.read_bytes() == b"abcde"
    assert seen_timeout == [120.0]  # transfer default
    assert not (tmp_path / "rec.mp4.part").exists()


def test_interrupted_download_leaves_no_partial_file(monkeypatch, tmp_path):
    monkeypatch.setattr(
        http_mod.urllib.request, "urlopen", lambda req, timeout=None: _Response(chunks=[b"ab", b"cd"], fail_after=1)
    )
    dest = tmp_path / "rec.mp4"
    with pytest.raises(TransportError):
        request_to_file("https://example.invalid/rec.mp4", dest, max_retries=0)
    assert list(tmp_path.iterdir()) == []


# ------------------------------------------------------------ threads


def test_parallel_scans_inherit_the_tool_timeout():
    from teams.client import TeamsClient

    convs = [{"id": f"19:{i}@thread.v2", "name": str(i)} for i in range(8)]
    with timeout_scope(42):
        results, errors = TeamsClient()._scan(convs, lambda conv: resolve_timeout("chat"))
    assert errors == [] and results == [42.0] * 8


# ------------------------------------------------------------ tools


@pytest.fixture
def server():
    mcp = MCPServer("timeouts")
    tools_mod.register_all(mcp)
    return mcp


@pytest.mark.anyio
async def test_timeout_parameter_is_documented_by_kind(server):
    tools = {t.name: t for t in await server.list_tools()}
    chat = tools["send_teams_message"].input_schema["properties"]["timeout_seconds"]["description"]
    transfer = tools["download_meeting_recordings"].input_schema["properties"]["timeout_seconds"]["description"]
    assert "not that the limit is too short" in chat
    assert "recordings" in transfer and "Raise it" in transfer
    for name in (
        "send_teams_message", "read_teams_chat", "react_to_teams_message", "edit_teams_message",
        "delete_teams_message", "list_teams_chats", "search_teams_chat_messages", "download_message_images",
        "download_chat_attachments", "download_sharepoint_link", "upload_sharepoint_file", "replace_sharepoint_file",
        "sync_folder_to_sharepoint", "download_meeting_recordings", "read_sharepoint_sheet",
        "update_sharepoint_sheet", "add_sharepoint_docx_comments",
    ):
        props = tools[name].input_schema["properties"]
        assert "timeout_seconds" in props, name
        assert "timeout_seconds" not in tools[name].input_schema.get("required", []), name


@pytest.mark.anyio
async def test_transfer_tool_timeout_reaches_the_http_layer(server, seen, monkeypatch):
    class FakeSharePoint:
        def download_link(self, url_or_guid, target_dir=""):
            with kind_scope("transfer"):
                http_mod.request_bytes("https://example.invalid/file.zip")
            return "ok"

    monkeypatch.setattr(tools_mod, "sp", lambda: FakeSharePoint())
    await server.call_tool("download_sharepoint_link", {"url_or_guid": "x", "timeout_seconds": 300})
    await server.call_tool("download_sharepoint_link", {"url_or_guid": "x"})
    await server.call_tool("download_sharepoint_link", {"url_or_guid": "x", "timeout_seconds": 9999})
    assert seen == [300.0, 120.0, 600.0]


@pytest.mark.anyio
async def test_chat_tool_timeout_applies_to_each_endpoint_attempt(server, monkeypatch, identity):
    from common.errors import ConnectError
    from teams.client import TeamsClient

    client = TeamsClient()
    monkeypatch.setattr(client, "_auth", lambda: {"region": "apac", "token": "t", "identity": identity})
    monkeypatch.setattr(tools_mod, "teams", lambda: client)
    attempts: list[tuple[str, float]] = []
    failed = []

    def fake_request(url, **kwargs):
        attempts.append((url.split("/")[2], kwargs["timeout"]))
        if not failed:
            failed.append(url)
            raise ConnectError("down")
        return 200, b'{"conversations": []}', {}

    monkeypatch.setattr("teams.client.request", fake_request)
    await server.call_tool("list_teams_chats", {"timeout_seconds": 42})
    assert attempts == [("teams.cloud.microsoft", 42.0), ("teams.microsoft.com", 42.0)]

    attempts.clear()
    client._conv_cache = None
    await server.call_tool("list_teams_chats", {})
    assert [t for _h, t in attempts] == [10.0]


def test_errors_keep_their_type_through_the_scope():
    @tools_mod._actionable
    def tool(timeout_seconds: float | None = None):
        raise Mcp365Error("hỏng", "sửa")

    from mcp.server.mcpserver.exceptions import ToolError

    with pytest.raises(ToolError, match="hỏng"):
        tool(timeout_seconds=5)
