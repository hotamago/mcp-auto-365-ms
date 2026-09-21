"""The tool surface registers cleanly and errors keep their remediation."""

from __future__ import annotations

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import tools as tools_mod
from common.errors import AuthExpiredError

EXPECTED_TOOLS = {
    # SharePoint / OneDrive
    "search_sharepoint_files",
    "read_sharepoint_link",
    "download_sharepoint_link",
    "upload_sharepoint_file",
    "replace_sharepoint_file",
    "compare_sharepoint_versions",
    "sync_folder_to_sharepoint",
    "download_meeting_recordings",
    "read_sharepoint_sheet",
    "update_sharepoint_sheet",
    "add_sharepoint_docx_comments",
    "get_word_companion_status",
    # Teams
    "list_teams_chats",
    "read_teams_chat",
    "get_recent_team_messages",
    "get_my_mentions",
    "get_new_mentions_since",
    "search_teams_chat_messages",
    "send_teams_message",
    "reply_to_channel_thread",
    "edit_teams_message",
    "delete_teams_message",
    "react_to_teams_message",
    "download_chat_attachments",
    "get_calendar_today",
    # Outlook mail
    "list_emails",
    "read_email",
    "send_email",
    # Cross-cutting
    "check_365_connection",
    "extract_action_items",
    "get_daily_briefing",
}


@pytest.fixture
def server():
    mcp = MCPServer("test-server")
    tools_mod.register_all(mcp)
    return mcp


@pytest.mark.anyio
async def test_all_expected_tools_are_registered(server):
    names = {t.name for t in await server.list_tools()}
    assert EXPECTED_TOOLS <= names


@pytest.mark.anyio
async def test_every_tool_has_a_description(server):
    for tool in await server.list_tools():
        assert tool.description, f"{tool.name} thiếu docstring"


@pytest.mark.anyio
async def test_prompts_and_resources_are_registered(server):
    assert {p.name for p in await server.list_prompts()} >= {
        "summarize_chat_thread",
        "draft_standup",
        "triage_mentions",
    }
    assert {str(r.uri) for r in await server.list_resources()} >= {"m365://health"}


def test_registering_twice_does_not_crash():
    """Standalone servers mount overlapping subsets of the same registry."""
    mcp = MCPServer("t")
    tools_mod.register_sharepoint_tools(mcp)
    tools_mod.register_resources(mcp)
    assert mcp is not None


def test_actionable_converts_typed_error_to_tool_error():
    """Only a ToolError message survives; anything else is masked by the SDK."""

    @tools_mod._actionable
    def boom():
        raise AuthExpiredError("Phiên hết hạn.", "Chạy az login.")

    with pytest.raises(ToolError) as excinfo:
        boom()
    assert "Phiên hết hạn." in str(excinfo.value)
    assert "Chạy az login." in str(excinfo.value)


def test_actionable_preserves_signature_and_doc():
    @tools_mod._actionable
    def sample(a: int, b: str = "x") -> str:
        """Docstring kept."""
        return f"{a}{b}"

    import inspect

    assert sample.__doc__ == "Docstring kept."
    assert list(inspect.signature(sample).parameters) == ["a", "b"]
    assert sample(1) == "1x"


def test_errors_note_renders_partial_failures():
    note = tools_mod._errors_note(["chat A: 429", "chat B: timeout"])
    assert "chưa đầy đủ" in note and "chat A" in note


def test_errors_note_empty_when_no_failures():
    assert tools_mod._errors_note([]) == ""
