"""Every outbound tool demands an explicit, required user confirmation."""

from __future__ import annotations

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import tools as tools_mod
from common import approval
from common.errors import ApprovalRequiredError

#: Every tool that sends, edits, deletes or overwrites something.
OUTBOUND_TOOLS = {
    "send_teams_message",
    "reply_to_channel_thread",
    "edit_teams_message",
    "delete_teams_message",
    "send_email",
    "upload_sharepoint_file",
    "replace_sharepoint_file",
    "update_sharepoint_sheet",
    "add_sharepoint_docx_comments",
    "sync_folder_to_sharepoint",
}


def test_confirmed_action_passes():
    approval.require_confirm(True, "Gửi", "đích", "nội dung")


@pytest.mark.parametrize("value", [False, None, "true", 1])
def test_anything_but_literal_true_is_refused(value):
    """A truthy string or 1 is not the user saying yes."""
    with pytest.raises(ApprovalRequiredError):
        approval.require_confirm(value, "Gửi", "đích", "nội dung")


def test_refusal_carries_the_draft_to_show_the_user():
    with pytest.raises(ApprovalRequiredError) as excinfo:
        approval.require_confirm(False, "Gửi tin nhắn Teams", "1:1 Chat (Hiển)", "Dạ em nghĩ không cần model")
    assert "CHƯA GỬI" in excinfo.value.message
    assert "1:1 Chat (Hiển)" in excinfo.value.message
    assert "Dạ em nghĩ không cần model" in excinfo.value.message
    assert "is_user_confirm=true" in excinfo.value.remediation


@pytest.fixture
def schemas():
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)
    import anyio

    return {t.name: t.input_schema for t in anyio.run(mcp.list_tools)}


def test_every_outbound_tool_requires_is_user_confirm(schemas):
    """Required in the schema, with the ask-the-user rule as its description."""
    for name in OUTBOUND_TOOLS:
        schema = schemas[name]
        assert "is_user_confirm" in schema.get("required", []), f"{name}: is_user_confirm không bắt buộc"
        description = schema["properties"]["is_user_confirm"].get("description", "")
        assert "hỏi ý kiến người dùng" in description, f"{name}: thiếu mô tả phải hỏi user"


def test_no_read_only_tool_asks_for_confirmation(schemas):
    for name, schema in schemas.items():
        if name not in OUTBOUND_TOOLS:
            assert "is_user_confirm" not in schema.get("properties", {}), name


@pytest.mark.anyio
async def test_unapproved_email_returns_exact_draft_before_touching_mail_client(monkeypatch):
    def unexpected_mail_access():
        raise AssertionError("mail client must not be touched before approval")

    monkeypatch.setattr(tools_mod, "outlook", unexpected_mail_access)
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)

    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "send_email",
            {
                "to": ["to@example.com"],
                "cc": ["cc@example.com"],
                "bcc": ["bcc@example.com"],
                "subject": "Exact subject",
                "body": "Exact body",
                "is_user_confirm": False,
            },
        )

    refusal = str(excinfo.value)
    assert "CHƯA GỬI" in refusal
    assert "To: to@example.com" in refusal
    assert "CC: cc@example.com" in refusal
    assert "BCC: bcc@example.com" in refusal
    assert "**Subject:** Exact subject" in refusal
    assert "Exact body" in refusal
