"""The outbound-approval gate: staging, confirmation and destination policy."""

from __future__ import annotations

import pytest

from common import approval
from common.config import reset_config_cache
from common.errors import ApprovalRequiredError, PendingActionError


@pytest.fixture(autouse=True)
def _clean_pending():
    approval.reset()
    yield
    approval.reset()


DIRECT = {"id": "19:aaa@unq.gbl.spaces", "name": "1:1 Chat (Nam Sơn)", "type": "DirectChat"}
GROUP = {"id": "19:bbb@thread.v2", "name": "Proactive Agent Feature Dev", "type": "GroupChat"}
CHANNEL = {"id": "19:ccc@thread.tacv2", "name": "[VF] #General", "type": "Channel"}
UNKNOWN_GROUP = {"id": "19:ddd@thread.v2", "name": "19:ddd@thread.v2", "type": "Unknown"}
NOTES = {"id": "48:notes", "name": "Notes", "type": "DirectChat"}


# ------------------------------------------------------- destination policy


@pytest.mark.parametrize("conv", [GROUP, CHANNEL, UNKNOWN_GROUP])
def test_multi_person_destinations_are_blocked_by_default(conv):
    with pytest.raises(ApprovalRequiredError) as excinfo:
        approval.guard_destination(conv, "gửi tin nhắn")
    assert "MCP365_ALLOW_GROUP_SENDS" in excinfo.value.remediation


def test_unknown_type_falls_back_to_the_id_and_stays_blocked():
    """A raw thread id not in the cache must not be mistaken for a 1:1."""
    assert approval.is_multi_person(UNKNOWN_GROUP) is True


def test_direct_chat_is_not_blocked():
    approval.guard_destination(DIRECT, "gửi tin nhắn")


def test_group_send_allowed_once_the_human_opts_in(monkeypatch):
    monkeypatch.setenv("MCP365_ALLOW_GROUP_SENDS", "1")
    reset_config_cache()
    approval.guard_destination(GROUP, "gửi tin nhắn")


def test_bool_env_ignores_the_word_false(monkeypatch):
    """``bool("false")`` is True - that mistake would unlock the gate."""
    monkeypatch.setenv("MCP365_ALLOW_GROUP_SENDS", "false")
    reset_config_cache()
    with pytest.raises(ApprovalRequiredError):
        approval.guard_destination(GROUP, "gửi tin nhắn")


# -------------------------------------------------------------- staging


def test_stage_does_not_execute():
    fired = []
    out = approval.stage("Gửi tin nhắn Teams", "1:1 Nam Sơn", "xin chào", lambda: fired.append(1) or "sent")
    assert fired == []
    assert "CHƯA GỬI" in out
    assert "xin chào" in out
    assert len(approval.pending_rows()) == 1


def test_confirm_executes_once():
    fired = []
    approval.stage("Gửi", "đích", "nội dung", lambda: fired.append(1) or "sent")
    token = approval.pending_rows()[0]["token"]

    assert approval.confirm(token) == "sent"
    assert fired == [1]

    with pytest.raises(PendingActionError):
        approval.confirm(token)


def test_confirm_tolerates_quoted_token():
    approval.stage("Gửi", "đích", "nội dung", lambda: "sent")
    token = approval.pending_rows()[0]["token"]
    assert approval.confirm(f'`"{token}"`') == "sent"


def test_cancel_discards_without_executing():
    fired = []
    approval.stage("Gửi", "đích", "nội dung", lambda: fired.append(1))
    token = approval.pending_rows()[0]["token"]

    assert "Đã huỷ" in approval.cancel(token)
    assert fired == []
    assert approval.pending_rows() == []


def test_unknown_token_is_actionable():
    with pytest.raises(PendingActionError) as excinfo:
        approval.confirm("deadbeef")
    assert "list_pending_actions" in excinfo.value.remediation


def test_expired_drafts_are_purged(monkeypatch):
    monkeypatch.setenv("MCP365_PENDING_TTL", "0")
    reset_config_cache()
    approval.stage("Gửi", "đích", "nội dung", lambda: "sent")
    assert approval.pending_rows() == []


def test_approval_can_be_switched_off_entirely(monkeypatch):
    monkeypatch.setenv("MCP365_REQUIRE_APPROVAL", "0")
    reset_config_cache()
    assert approval.stage("Gửi", "đích", "nội dung", lambda: "sent") == "sent"


def test_self_chat_detection():
    assert approval.is_self_chat(NOTES["id"]) is True
    assert approval.is_self_chat(DIRECT["id"]) is False
