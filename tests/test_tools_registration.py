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
    "delete_sharepoint_item",
    "compare_sharepoint_versions",
    "sync_folder_to_sharepoint",
    "download_meeting_recordings",
    "read_sharepoint_sheet",
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
    "download_message_images",
    "find_user",
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


#: Tools that edited an existing file in place, removed on 26/09: writes into a
#: file others had open were refused (423/412) or re-uploaded the whole file.
REMOVED_TOOLS = {"update_sharepoint_sheet", "add_sharepoint_docx_comments", "replace_sharepoint_file"}


@pytest.mark.anyio
async def test_all_expected_tools_are_registered(server):
    names = {t.name for t in await server.list_tools()}
    assert EXPECTED_TOOLS <= names
    assert not REMOVED_TOOLS & names


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


def test_actionable_turns_an_unexpected_crash_into_a_described_tool_error():
    """A bare KeyError used to reach the agent as "Error executing tool send_teams_message"."""

    @tools_mod._actionable
    def boom():
        return {}["conversation_id"]

    with pytest.raises(ToolError) as excinfo:
        boom()
    text = str(excinfo.value)
    assert "KeyError" in text and "conversation_id" in text
    # Where it broke: the raising function, not _actionable's own wrapper.
    assert "test_tools_registration.py:" in text and "(boom)" in text
    assert "wrapper" not in text
    assert isinstance(excinfo.value.__cause__, KeyError)


def test_actionable_names_the_src_module_for_crashes_inside_the_server():
    from sharepoint.client import SharePointClient

    client = SharePointClient()
    client._drive_cache["web:drive"] = "https://t.sharepoint.com/sites/X/Shared Documents"

    @tools_mod._actionable
    def boom():
        return client._item_file_url("drive", {"parentReference": {}})  # item without "name"

    with pytest.raises(ToolError) as excinfo:
        boom()
    # client.py alone would be ambiguous (sharepoint/teams/outlook).
    assert "sharepoint/client.py:" in str(excinfo.value)


def test_actionable_passes_tool_errors_through_untouched():
    @tools_mod._actionable
    def boom():
        raise ToolError("đã rõ ràng")

    with pytest.raises(ToolError, match="^đã rõ ràng$"):
        boom()
