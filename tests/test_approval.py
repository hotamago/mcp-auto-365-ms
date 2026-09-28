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
    "react_to_teams_message",
    "send_email",
    "reply_email",
    "upload_sharepoint_file",
    "delete_sharepoint_item",
    "sync_folder_to_sharepoint",
    "share_file_onedrive",
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


@pytest.mark.anyio
async def test_unapproved_reaction_returns_exact_target_before_mutation(monkeypatch):
    class FakeTeams:
        def find_conversation(self, _identifier):
            return {"id": "19:dev@thread.v2", "name": "Dev team"}

        def react_to_message(self, *_args, **_kwargs):
            raise AssertionError("reaction must not be changed before approval")

    monkeypatch.setattr(tools_mod, "teams", lambda: FakeTeams())
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)

    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "react_to_teams_message",
            {
                "chat_name_or_id": "Dev team",
                "message_id": "1789977000123",
                "reaction": "like",
                "remove": False,
                "is_user_confirm": False,
            },
        )

    refusal = str(excinfo.value)
    assert "CHƯA GỬI" in refusal
    assert "Dev team" in refusal
    assert "1789977000123" in refusal
    assert "👍" in refusal


@pytest.mark.anyio
async def test_unapproved_onedrive_share_uploads_nothing_and_shows_who_can_open(monkeypatch):
    def unexpected_sharepoint_access():
        raise AssertionError("SharePoint client must not be touched before approval")

    monkeypatch.setattr(tools_mod, "sp", unexpected_sharepoint_access)
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)
    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool("share_file_onedrive", {"local_file_path": "/tmp/bao-cao.md", "is_user_confirm": False})
    refusal = str(excinfo.value)
    assert "OneDrive của bạn › `Shared from MCP`" in refusal
    assert "`/tmp/bao-cao.md` → `bao-cao.md`" in refusal
    assert "mọi người trong tổ chức có link đều **xem** được" in refusal


@pytest.mark.anyio
async def test_unapproved_attachment_preview_says_where_the_file_goes_and_who_can_open_it(monkeypatch):
    class FakeTeams:
        def find_conversation(self, identifier):
            if identifier == "General":
                return {"id": "19:c@thread.tacv2", "name": "[VF] #General", "type": "Channel"}
            return {"id": "19:dev@thread.v2", "name": "Dev team", "type": "GroupChat"}

        def send_message(self, *_args, **_kwargs):
            raise AssertionError("nothing may be uploaded or sent before approval")

    monkeypatch.setattr(tools_mod, "teams", lambda: FakeTeams())
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)

    async def refusal(**args):
        with pytest.raises(ToolError) as excinfo:
            await mcp.call_tool("send_teams_message", {"message": "đây", "file_path": "/tmp/a.xlsx",
                                                       "is_user_confirm": False, **args})
        return str(excinfo.value)

    text = await refusal(chat_name_or_id="Dev team")
    assert "OneDrive của bạn › `Microsoft Teams Chat Files`" in text
    assert "chỉ các thành viên của chat này" in text and "không gửi email mời" in text
    assert "mọi người trong tổ chức có link" in await refusal(chat_name_or_id="Dev team", share_scope="organization")
    assert "SharePoint ›" in await refusal(chat_name_or_id="General")
    with pytest.raises(ToolError, match="share_scope"):
        await mcp.call_tool("send_teams_message", {"chat_name_or_id": "Dev team", "message": "x",
                                                   "is_user_confirm": False, "share_scope": "all"})


@pytest.mark.anyio
async def test_unapproved_email_with_files_lists_name_size_and_how_each_is_attached(monkeypatch, tmp_path):
    """Chỉ dựa vào kích thước file cục bộ: bản nháp có đủ tên, cỡ, cách gắn mà không gọi mạng."""
    from outlook import client as mail_mod

    def unexpected_mail_access():
        raise AssertionError("mail client must not be touched before approval")

    monkeypatch.setattr(tools_mod, "outlook", unexpected_mail_access)
    monkeypatch.setattr(mail_mod, "DIRECT_ATTACH_LIMIT", 10)
    monkeypatch.setattr(mail_mod, "SESSION_ATTACH_LIMIT", 100)
    small, medium, big = tmp_path / "a.txt", tmp_path / "b.pdf", tmp_path / "c.zip"
    small.write_bytes(b"x" * 5)
    medium.write_bytes(b"x" * 50)
    big.write_bytes(b"x" * 500)
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)

    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "send_email",
            {
                "to": ["to@example.com"],
                "subject": "S",
                "body": "**B**",
                "body_format": "markdown",
                "attachments": [str(small), str(medium), str(big)],
                "is_user_confirm": False,
            },
        )
    refusal = str(excinfo.value)
    assert "CHƯA GỬI" in refusal and "markdown → HTML" in refusal and "**B**" in refusal
    assert "`a.txt` · 5 B · gắn trực tiếp" in refusal
    assert "`b.pdf` · 50 B · gắn trực tiếp (upload session" in refusal
    assert "`c.zip` · 500 B · link: tải lên OneDrive của bạn › `Attachments`" in refusal
    assert "chỉ người nhận To/CC/BCC mở được" in refusal


@pytest.mark.anyio
async def test_unapproved_email_with_missing_file_fails_before_the_draft(monkeypatch, tmp_path):
    monkeypatch.setattr(tools_mod, "outlook", lambda: (_ for _ in ()).throw(AssertionError("no mail access")))
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)
    with pytest.raises(ToolError, match="Không tìm thấy file đính kèm"):
        await mcp.call_tool(
            "send_email",
            {"to": ["a@example.com"], "subject": "S", "body": "B", "attachments": [str(tmp_path / "nope.txt")],
             "is_user_confirm": False},
        )


@pytest.mark.anyio
async def test_unapproved_reply_shows_recipients_subject_and_original_and_writes_nothing(monkeypatch):
    class FakeOutlook:
        def prepare_reply(self, message_id, mode, to=None, cc=None, bcc=None):
            assert (message_id, mode, cc) == ("M1", "reply_all", ["boss@example.com"])
            return {
                "original": {"id": "M1", "subject": "Kế hoạch", "from": "Nam <nam@example.com>",
                             "received": "2026-09-28T01:00:00Z", "has_attachments": True},
                "mode": mode,
                "to": ["nam@example.com", "hien@example.com"],
                "cc": ["boss@example.com"],
                "bcc": [],
                "subject": "RE: Kế hoạch",
            }

        def compose_and_send(self, **_kwargs):
            raise AssertionError("nothing may be drafted or sent before approval")

    monkeypatch.setattr(tools_mod, "outlook", lambda: FakeOutlook())
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)
    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(
            "reply_email",
            {"message_id": "M1", "mode": "reply_all", "body": "Ok anh", "cc": ["boss@example.com"],
             "is_user_confirm": False},
        )
    refusal = str(excinfo.value)
    assert "Trả lời tất cả email cần người dùng duyệt" in refusal
    assert "To: nam@example.com, hien@example.com" in refusal and "CC: boss@example.com" in refusal
    assert "**Subject:** RE: Kế hoạch" in refusal
    assert "Kế hoạch — từ Nam <nam@example.com>" in refusal and "Message ID `M1`" in refusal
    assert "Ok anh" in refusal
    assert "file đính kèm của mail gốc" not in refusal  # chỉ forward mới mang file gốc theo
