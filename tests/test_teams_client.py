"""Teams client behaviour that can be verified without a network."""

from __future__ import annotations

from datetime import datetime, timedelta

import pytest

from teams.client import TeamsClient, clean_teams_html, parse_since, text_to_teams_html

CONVERSATIONS = [
    {"id": "19:abc@thread.v2", "name": "Dev team", "type": "GroupChat", "last_activity": "", "last_sender": "", "last_message": ""},
    {"id": "19:xyz@thread.tacv2", "name": "[VF] #General", "type": "Channel", "last_activity": "", "last_sender": "", "last_message": ""},
]


@pytest.fixture
def client(monkeypatch):
    c = TeamsClient()
    calls = {"n": 0}

    def fake_list(page_size=50, filter_keyword="", use_cache=True):
        calls["n"] += 1
        return list(CONVERSATIONS)

    monkeypatch.setattr(c, "list_conversations", fake_list)
    c._list_calls = calls
    return c


# ------------------------------------------------------------ HTML handling


def test_outgoing_html_is_escaped():
    """`a < b` must not be swallowed by Teams as a bogus tag."""
    assert text_to_teams_html("a < b & c > d") == "<p>a &lt; b &amp; c &gt; d</p>"


def test_markdown_is_still_converted():
    out = text_to_teams_html("**bold** and `code`")
    assert "<b>bold</b>" in out and "<code>code</code>" in out


def test_incoming_entities_are_decoded():
    # The old hand-rolled table only knew five entities; &eacute; leaked through.
    assert clean_teams_html("<p>caf&eacute; &amp; 5 &lt; 6</p>") == "café & 5 < 6"


def test_mention_span_becomes_readable_text():
    html = '<span itemtype="http://schema.skype.com/Mention" itemid="0">Sơn</span> xem nhé'
    assert clean_teams_html(html).startswith("@Sơn")


# ------------------------------------------------------------------- since


def test_since_today_uses_local_midnight():
    """A UTC-midnight 'today' silently dropped 00:00-07:00 for a UTC+7 user."""
    result = parse_since("today")
    local_now = datetime.now().astimezone()
    assert result.hour == 0 and result.minute == 0
    assert result.utcoffset() == local_now.utcoffset()
    assert result.date() == local_now.date()


def test_since_relative_forms():
    now = datetime.now().astimezone()
    assert abs((now - parse_since("6h")) - timedelta(hours=6)) < timedelta(seconds=5)
    assert abs((now - parse_since("2d")) - timedelta(days=2)) < timedelta(seconds=5)


def test_since_iso_date_is_local():
    assert parse_since("2026-09-18").utcoffset() == datetime.now().astimezone().utcoffset()


def test_since_invalid_returns_none():
    assert parse_since("hôm kia kìa") is None
    assert parse_since("") is None


# ------------------------------------------------- conversation resolution


def test_id_shaped_identifier_skips_the_listing(client):
    """Regression guard for the N+1: an ID must not trigger a full listing."""
    conv = client.find_conversation("19:someid@thread.v2")
    assert conv["id"] == "19:someid@thread.v2"
    assert client._list_calls["n"] == 0


def test_self_aliases_resolve_without_network(client):
    for alias in ("me", "self", "48:notes", "Notes"):
        assert client.find_conversation(alias)["id"] == "48:notes"
    assert client._list_calls["n"] == 0


def test_name_lookup_uses_the_listing(client):
    assert client.find_conversation("Dev team")["id"] == "19:abc@thread.v2"
    assert client._list_calls["n"] == 1


def test_partial_name_match(client):
    assert client.find_conversation("dev")["id"] == "19:abc@thread.v2"


def test_unknown_name_raises_with_suggestions(client):
    from common.errors import ConversationNotFoundError

    with pytest.raises(ConversationNotFoundError) as excinfo:
        client.find_conversation("khong-ton-tai")
    assert "Dev team" in excinfo.value.remediation


def test_channel_thread_reply_rejects_plain_group_chat(client):
    from common.errors import UnsupportedOperationError

    with pytest.raises(UnsupportedOperationError):
        client.reply_to_channel_thread("Dev team", "123", "hi")


# ------------------------------------------------------- partial failures


def test_scan_reports_failures_instead_of_swallowing(client):
    from common.errors import Mcp365Error

    def worker(conv):
        if conv["type"] == "Channel":
            raise Mcp365Error("bị giới hạn tần suất", "thử lại sau")
        return {"ok": conv["name"]}

    results, errors = client._scan(list(CONVERSATIONS), worker)
    assert len(results) == 1
    assert len(errors) == 1 and "[VF] #General" in errors[0]


# --------------------------------------------------- sent-message id lookup


def test_sent_id_resolved_by_client_message_id(client, monkeypatch):
    """The send endpoint returns no id, so it is recovered via clientmessageid."""
    monkeypatch.setattr(
        client,
        "get_messages",
        lambda *a, **k: {
            "messages": [
                {"id": "111", "raw": {"clientmessageid": "999"}},
                {"id": "222", "raw": {"clientmessageid": "1234"}},
            ]
        },
    )
    assert client._resolve_sent_id("48:notes", "1234", {}) == "222"


def test_sent_id_falls_back_to_arrival_time(client, monkeypatch):
    monkeypatch.setattr(client, "get_messages", lambda *a, **k: {"messages": []})
    assert client._resolve_sent_id("48:notes", "1234", {"OriginalArrivalTime": 1789926111943}) == "1789926111943"


def test_sent_id_survives_lookup_failure(client, monkeypatch):
    from common.errors import Mcp365Error

    def boom(*a, **k):
        raise Mcp365Error("mạng lỗi", "thử lại")

    monkeypatch.setattr(client, "get_messages", boom)
    assert client._resolve_sent_id("48:notes", "1234", {"OriginalArrivalTime": 42}) == "42"
