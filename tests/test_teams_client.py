"""Teams client behaviour that can be verified without a network."""

from __future__ import annotations

from datetime import datetime, timedelta

import pytest

from teams.client import TeamsClient, clean_teams_html, parse_since, text_to_teams_html

CONVERSATIONS = [
    {
        "id": "19:abc@thread.v2",
        "name": "Dev team",
        "type": "GroupChat",
        "last_activity": "",
        "last_sender": "",
        "last_message": "",
    },
    {
        "id": "19:xyz@thread.tacv2",
        "name": "[VF] #General",
        "type": "Channel",
        "last_activity": "",
        "last_sender": "",
        "last_message": "",
    },
]


@pytest.fixture
def client(monkeypatch, identity):
    c = TeamsClient()
    calls = {"n": 0}

    def fake_list(page_size=50, filter_keyword="", use_cache=True):
        calls["n"] += 1
        return list(CONVERSATIONS)

    monkeypatch.setattr(c, "list_conversations", fake_list)
    monkeypatch.setattr(type(c), "identity", property(lambda self: identity))
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


def test_failed_attachment_upload_says_the_message_was_not_sent(client, monkeypatch, tmp_path):
    from common.errors import Mcp365Error
    from sharepoint.client import SharePointClient

    local = tmp_path / "spec.pdf"
    local.write_bytes(b"%PDF")

    def refuse(self, *args, **kwargs):
        raise Mcp365Error("Cả 2 kênh SharePoint đều lỗi khi upload 'spec.pdf'.", "Kiểm tra quyền.")

    monkeypatch.setattr(client, "_auth", lambda: {"base_url": "https://chat.example/v1", "token": "t"})
    monkeypatch.setattr(SharePointClient, "upload_file", refuse)
    monkeypatch.setattr("teams.client.request_json", lambda *a, **k: pytest.fail("nothing may be sent"))

    with pytest.raises(Mcp365Error) as excinfo:
        client.send_message("Dev team", "hi", file_path=str(local))
    assert "CHƯA được gửi" in excinfo.value.message and "Cả 2 kênh SharePoint" in excinfo.value.message
    assert excinfo.value.remediation == "Kiểm tra quyền."


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


# -------------------------------------------------------- message reactions


def test_add_reaction_uses_chat_service_emotions_property(client, monkeypatch):
    import json

    captured = {}
    monkeypatch.setattr(
        client, "_auth", lambda: {"base_url": "https://emea.ng.msg.teams.microsoft.com/v1", "token": "t"}
    )
    monkeypatch.setattr("teams.client.time.time", lambda: 1789977600.123)

    def fake_request(url, **kwargs):
        captured["url"] = url
        captured.update(kwargs)
        return 200, b"", {}

    monkeypatch.setattr("teams.client.request", fake_request)

    result = client.react_to_message("Dev team", "1789977000123", "👍")

    assert captured["url"] == (
        "https://emea.ng.msg.teams.microsoft.com/v1/users/ME/conversations/"
        "19%3Aabc%40thread.v2/messages/1789977000123/properties?name=emotions"
    )
    assert captured["method"] == "PUT"
    assert captured["headers"]["x-ms-client-caller"] == "updateMessageReactionAdd"
    assert json.loads(captured["data"]) == {"emotions": '{"key":"like","value":1789977600123}'}
    assert result["status"] == "REACTED"
    assert result["reaction"] == "like"
    assert result["emoji"] == "👍"


def test_remove_reaction_uses_delete_without_timestamp(client, monkeypatch):
    import json

    captured = {}
    monkeypatch.setattr(client, "_auth", lambda: {"base_url": "https://chat.example/v1", "token": "t"})

    def fake_request(url, **kwargs):
        captured.update(kwargs)
        return 200, b"", {}

    monkeypatch.setattr("teams.client.request", fake_request)

    result = client.react_to_message("19:abc@thread.v2", "123", "surprise", remove=True)

    assert captured["method"] == "DELETE"
    assert captured["headers"]["x-ms-client-caller"] == "updateMessageReactionRemove"
    assert json.loads(captured["data"]) == {"emotions": '{"key":"surprised"}'}
    assert result["status"] == "REMOVED"
    assert result["reaction"] == "surprised"


def test_invalid_reaction_is_rejected_before_network(client):
    from common.errors import ConfigError

    with pytest.raises(ConfigError, match="Reaction Teams không hợp lệ"):
        client.react_to_message("Dev team", "123", "thumbs-up")


# ------------------------------------------ keyword filter & diacritics


def _cached_client(identity, monkeypatch, conversations):
    """A client whose conversation cache is pre-warmed, so no network is used."""
    import time as _time

    c = TeamsClient()
    monkeypatch.setattr(type(c), "identity", property(lambda self: identity))
    c._conv_cache = list(conversations)
    c._conv_cache_at = _time.time()
    return c


MANY = [
    {
        "id": f"19:filler{i}@thread.v2",
        "name": f"Nhóm {i}",
        "type": "GroupChat",
        "last_activity": "",
        "last_sender": "",
        "last_message": "",
    }
    for i in range(20)
] + [
    {
        "id": "19:namson@unq.gbl.spaces",
        "name": "1:1 Chat (Nguyễn Phan Nam Sơn)",
        "type": "DirectChat",
        "last_activity": "",
        "last_sender": "",
        "last_message": "",
    }
]


def test_fold_strips_vietnamese_diacritics():
    from teams.client import fold

    assert fold("Nguyễn Phan Nam Sơn") == "nguyen phan nam son"
    assert fold("Đỗ Văn Hoàng") == "do van hoang"


def test_keyword_filter_reaches_past_the_limit(identity, monkeypatch):
    """The match sits at index 20; asking for 5 rows must still find it.

    Truncating before filtering used to hide every match outside the first
    page - the reason a 1:1 chat 23 rows down was reported as non-existent.
    """
    c = _cached_client(identity, monkeypatch, MANY)
    found = c.list_conversations(page_size=5, filter_keyword="Nam Sơn")
    assert [x["id"] for x in found] == ["19:namson@unq.gbl.spaces"]


def test_keyword_filter_ignores_diacritics(identity, monkeypatch):
    c = _cached_client(identity, monkeypatch, MANY)
    assert len(c.list_conversations(page_size=50, filter_keyword="nam son")) == 1


def test_limit_still_applies_without_a_keyword(identity, monkeypatch):
    c = _cached_client(identity, monkeypatch, MANY)
    assert len(c.list_conversations(page_size=5)) == 5


def test_chat_type_filter(identity, monkeypatch):
    c = _cached_client(identity, monkeypatch, MANY)
    rows = c.list_conversations(page_size=50, chat_type="DirectChat")
    assert [r["type"] for r in rows] == ["DirectChat"]


def test_find_conversation_matches_without_diacritics(identity, monkeypatch):
    c = _cached_client(identity, monkeypatch, MANY)
    monkeypatch.setattr(c, "list_conversations", lambda **kw: list(MANY))
    assert c.find_conversation("nam son")["id"] == "19:namson@unq.gbl.spaces"


# ------------------------------------------------------------ attachments


def test_attachments_come_from_properties_files_not_the_body():
    """A file-only message has an empty body; the file is in properties.files."""
    import json as _json

    from teams.client import parse_attachments

    raw = {
        "content": "",
        "properties": {
            "files": _json.dumps(
                [
                    {
                        "fileName": "PSDK.zip",
                        "fileType": "zip",
                        "objectUrl": "https://t-my.sharepoint.com/personal/u/Documents/Microsoft Teams Chat Files/PSDK.zip",
                        "fileInfo": {"shareUrl": "https://t-my.sharepoint.com/:u:/g/personal/u/IQDZ"},
                    }
                ]
            )
        },
    }
    assert parse_attachments(raw) == [
        {
            "name": "PSDK.zip",
            "type": "zip",
            "url": "https://t-my.sharepoint.com/personal/u/Documents/Microsoft Teams Chat Files/PSDK.zip",
            "share_url": "https://t-my.sharepoint.com/:u:/g/personal/u/IQDZ",
        }
    ]


def test_attachments_tolerate_missing_or_malformed_payloads():
    from teams.client import parse_attachments

    assert parse_attachments({}) == []
    assert parse_attachments({"properties": {"files": "not json"}}) == []
    assert parse_attachments({"properties": {"files": "[]"}}) == []


# ------------------------------------------------------------ mentions


def test_mention_replaces_the_at_name_and_builds_properties():
    from teams.client import apply_mentions, text_to_teams_html

    people = [{"name": "Phạm Sỹ Hùng", "display_name": "Phạm Sỹ Hùng (VF-KPTX-VPTAITX)", "mri": "8:orgid:aaa"}]
    html, props = apply_mentions(text_to_teams_html("@Phạm Sỹ Hùng ơi, GPU dev có chưa ạ?"), people)
    assert html.startswith('<p><span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="0">')
    assert "@Phạm Sỹ Hùng" not in html
    assert props == [
        {
            "@type": "http://schema.skype.com/Mention",
            "itemid": "0",
            "mri": "8:orgid:aaa",
            "mentionType": "person",
            "displayName": "Phạm Sỹ Hùng (VF-KPTX-VPTAITX)",
        }
    ]


def test_person_not_written_in_text_is_tagged_up_front():
    """Asking to tag someone must never silently tag nobody."""
    from teams.client import apply_mentions, text_to_teams_html

    people = [
        {"name": "Hùng", "display_name": "Phạm Sỹ Hùng", "mri": "8:orgid:a"},
        {"name": "Hoàng", "display_name": "Đỗ Văn Hoàng", "mri": "8:orgid:b"},
    ]
    html, props = apply_mentions(text_to_teams_html("Anh ơi GPU dev có chưa?"), people)
    assert html.count("schema.skype.com/Mention") == 2
    assert html.index('itemid="0"') < html.index("Anh ơi")
    assert [p["itemid"] for p in props] == ["0", "1"]


def _history_client(identity, monkeypatch, messages):
    c = _cached_client(
        identity,
        monkeypatch,
        [
            {
                "id": "19:g@thread.v2",
                "name": "Dev team",
                "type": "GroupChat",
                "last_activity": "",
                "last_sender": "",
                "last_message": "",
            }
        ],
    )
    monkeypatch.setattr(c, "get_messages", lambda conv, limit=200: {"messages": messages})
    return c


HISTORY = [
    {"sender": "Phạm Sỹ Hùng (VF-KPTX-VPTAITX)", "sender_mri": "8:orgid:hung", "mentions": []},
    {
        "sender": "Trịnh Anh Tuấn (VF-KPTX-VPTAITX)",
        "sender_mri": "8:orgid:tuan",
        "mentions": [{"mri": "8:orgid:hoang", "displayName": "Đỗ Văn Hoàng (VF-KPTX-VPTAITX)"}],
    },
    {"sender": "Hoàng Cao Minh (VF-KPTX-VPTAITX)", "sender_mri": "8:orgid:minh", "mentions": []},
]


def test_resolve_finds_senders_and_mentioned_people_without_diacritics(identity, monkeypatch):
    c = _history_client(identity, monkeypatch, HISTORY)
    people = c.resolve_mentions("19:g@thread.v2", ["pham sy hung", "@Đỗ Văn Hoàng"])
    assert [(p["mri"], p["name"]) for p in people] == [
        ("8:orgid:hung", "pham sy hung"),
        ("8:orgid:hoang", "Đỗ Văn Hoàng"),
    ]


def test_resolve_refuses_ambiguous_names(identity, monkeypatch):
    from common.errors import Mcp365Error

    c = _history_client(identity, monkeypatch, HISTORY)
    with pytest.raises(Mcp365Error) as excinfo:
        c.resolve_mentions("19:g@thread.v2", ["Hoàng"])  # Đỗ Văn Hoàng and Hoàng Cao Minh
    assert "nhiều người" in excinfo.value.message


def test_resolve_unknown_person_is_actionable(identity, monkeypatch):
    from common.errors import ConversationNotFoundError

    c = _history_client(identity, monkeypatch, HISTORY)
    with pytest.raises(ConversationNotFoundError):
        c.resolve_mentions("19:g@thread.v2", ["Người Lạ"])


def test_shortened_tag_seen_first_does_not_hide_the_full_name(identity, monkeypatch):
    """Someone tagged "Hoàng" before Hoàng wrote under his full name."""
    history = [
        {"sender": "Trịnh Anh Tuấn", "sender_mri": "8:orgid:tuan",
         "mentions": [{"mri": "8:orgid:hoang", "displayName": "Hoàng"}]},
        {"sender": "Đỗ Văn Hoàng (VF-KPTX-VPTAITX)", "sender_mri": "8:orgid:hoang", "mentions": []},
    ]
    c = _history_client(identity, monkeypatch, history)
    [person] = c.resolve_mentions("19:g@thread.v2", ["Đỗ Văn Hoàng"])
    assert person["mri"] == "8:orgid:hoang"
    assert person["display_name"] == "Đỗ Văn Hoàng (VF-KPTX-VPTAITX)"


def test_fold_handles_both_capital_d_with_stroke_lookalikes():
    from teams.client import fold

    assert fold("Đỗ") == fold("Ðỗ") == "do"


def test_parse_inline_images():
    from teams.client import parse_inline_images

    html = (
        '<p>Hello <img itemscope="" itemtype="http://schema.skype.com/Emoji" src="https://cdn/smile.png"> '
        '<img src="https://as-api.asm.skype.com/v1/objects/0-jhb-d10-3e07e06ab1b3434a1c63cc82c1f10ff7/views/imgo" '
        'itemtype="http://schema.skype.com/AMSImage" width="200" alt="screenshot"> end</p>'
    )
    imgs = parse_inline_images(html)
    assert len(imgs) == 1
    assert imgs[0]["id"] == "0-jhb-d10-3e07e06ab1b3434a1c63cc82c1f10ff7"
    assert "0-jhb-d10-3e" in imgs[0]["name"]
    assert imgs[0]["type"] == "inline"


def test_clean_teams_html_preserves_image_markers():
    from teams.client import clean_teams_html

    html = (
        '<p>Xem log này:<br>'
        '<img src="https://as-api.asm.skype.com/v1/objects/123/views/imgo" itemtype="http://schema.skype.com/AMSImage">'
        ' <img itemtype="http://schema.skype.com/Emoji" src="smile.png"></p>'
    )
    cleaned = clean_teams_html(html)
    assert "Xem log này:" in cleaned
    assert "🖼️ [image]" in cleaned
    assert "smile" not in cleaned


def test_search_users_directory(client, monkeypatch):
    class FakeMailAuth:
        def get_token(self):
            return "fake-mail-token"

    monkeypatch.setattr("outlook.auth.MailAuthManager", FakeMailAuth)

    def fake_request_json(url, headers=None, context=""):
        assert "fake-mail-token" in headers["Authorization"]
        return {
            "value": [
                {
                    "Id": "de4cb212-40ef-4a20-81c4-246c0da5458e@tenant",
                    "DisplayName": "Trịnh Anh Tuấn (VF)",
                    "GivenName": "Tuấn",
                    "Surname": "Trịnh Anh",
                    "JobTitle": "Chuyên gia AI",
                    "Department": "AI Squad",
                    "UserPrincipalName": "tuanta81@example.com",
                    "ScoredEmailAddresses": [{"Address": "v.tuanta81@example.com"}],
                    "Phones": [{"Number": "0912345678"}],
                }
            ]
        }

    monkeypatch.setattr("teams.client.request_json", fake_request_json)

    results = client.search_users("tuanta81")
    assert len(results) == 1
    u = results[0]
    assert u["name"] == "Trịnh Anh Tuấn (VF)"
    assert u["email"] == "v.tuanta81@example.com"
    assert u["teams_mri"] == "8:orgid:de4cb212-40ef-4a20-81c4-246c0da5458e"
    assert "de4cb212-40ef-4a20-81c4-246c0da5458e" in u["direct_chat_id"]


def test_download_message_images(client, monkeypatch, tmp_path):
    monkeypatch.setattr(client, "find_conversation", lambda name: {"id": "conv1", "name": "General"})
    monkeypatch.setattr(
        client,
        "get_messages",
        lambda cid, limit=50: {
            "messages": [
                {
                    "id": "17899770001",
                    "sender": "Tuấn",
                    "timestamp": "2026-09-21 10:00:00",
                    "images": [
                        {
                            "id": "img123",
                            "name": "screenshot.png",
                            "url": "https://as-api.asm.skype.com/v1/objects/img123/views/imgo",
                            "type": "inline",
                        }
                    ],
                }
            ]
        },
    )

    def fake_download_image(url, path):
        from pathlib import Path
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(b"PNG_MOCK")
        return p

    monkeypatch.setattr(client, "download_image", fake_download_image)

    downloaded = client.download_message_images("General", target_dir=str(tmp_path))
    assert len(downloaded) == 1
    assert downloaded[0]["message_id"] == "17899770001"
    assert (tmp_path / downloaded[0]["name"]).is_file()


def test_resolve_mentions_falls_back_to_directory_search(identity, monkeypatch):
    c = _history_client(identity, monkeypatch, [])  # Empty history

    def fake_search_users(query, max_results=3):
        if "nguyen van a" in query.lower():
            return [{"name": "Nguyễn Văn A", "teams_mri": "8:orgid:user_a_guid"}]
        return []

    monkeypatch.setattr(c, "search_users", fake_search_users)
    [person] = c.resolve_mentions("19:g@thread.v2", ["Nguyen Van A"])
    assert person["mri"] == "8:orgid:user_a_guid"
    assert person["display_name"] == "Nguyễn Văn A"
