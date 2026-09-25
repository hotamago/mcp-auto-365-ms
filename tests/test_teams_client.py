"""Teams client behaviour that can be verified without a network."""

from __future__ import annotations

import json
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
    monkeypatch.setattr("teams.client.request", lambda *a, **k: pytest.fail("nothing may be sent"))

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
    monkeypatch.setattr(client, "_auth", lambda: {"region": "emea", "token": "t"})
    monkeypatch.setattr("teams.client.time.time", lambda: 1789977600.123)

    def fake_request(url, **kwargs):
        captured["url"] = url
        captured.update(kwargs)
        return 200, b"", {}

    monkeypatch.setattr("teams.client.request", fake_request)

    result = client.react_to_message("Dev team", "1789977000123", "👍")

    # First configured front door, in the token's region; the proxy gets its own Origin.
    assert captured["url"] == (
        "https://teams.cloud.microsoft/api/chatsvc/emea/v1/users/ME/conversations/"
        "19%3Aabc%40thread.v2/messages/1789977000123/properties?name=emotions"
    )
    assert captured["headers"]["Origin"] == "https://teams.cloud.microsoft"
    assert captured["max_retries"] == 0
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


def _history_client(identity, monkeypatch, messages, members=()):
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
    monkeypatch.setattr(c, "get_members", lambda conv_id: set(members))
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


HOANG_MRI = "8:orgid:c254a4f9-539c-42e0-94a9-bc6ea1753622"
OTHER_MRI = "8:orgid:11111111-2222-3333-4444-555555555555"
HOANG_DIR = {
    "name": "Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM)",
    "email": "v.hoangnh21@vinfast.vn",
    "all_emails": ["v.hoangnh21@vinfast.vn"],
    "upn": "hoangnh21@vingroup.net",
    "teams_mri": HOANG_MRI,
}
OTHER_DIR = {
    "name": "Nguyễn Huy Hoàng (VF-KPTX-TKVPTLK)",
    "email": "v.hoangnh99@vinfast.vn",
    "all_emails": ["v.hoangnh99@vinfast.vn"],
    "upn": "hoangnh99@vingroup.net",
    "teams_mri": OTHER_MRI,
}


def _fake_directory(people):
    """People API stand-in: misses full addresses, as the real one does; matches names/local parts."""
    from teams.client import fold

    def search(query, max_results=3):
        q = fold(query)
        if "@" in q:
            return []
        return [p for p in people if q in fold(p["name"]) or any(q in a.split("@")[0] for a in [p["upn"], p["email"]])]

    return search


def test_moved_org_unit_is_tagged_under_current_name(identity, monkeypatch):
    """23/09 bug: same MRI, old unit first in history -> tag showed the old unit."""
    history = [
        {"sender": "Nguyễn Huy Hoàng (VF-KPTX-TKVPTLK)", "sender_mri": HOANG_MRI, "mentions": []},
        {"sender": "Trịnh Anh Tuấn", "sender_mri": "8:orgid:tuan",
         "mentions": [{"mri": HOANG_MRI, "displayName": "Nguyễn"}, {"mri": HOANG_MRI, "displayName": "Hoàng"}]},
        {"sender": "Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM)", "sender_mri": HOANG_MRI, "mentions": []},
    ]
    c = _history_client(identity, monkeypatch, history)
    for asked in ["Nguyễn Huy Hoàng", "Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM)", HOANG_MRI]:
        [person] = c.resolve_mentions("19:g@thread.v2", [asked])
        assert person["mri"] == HOANG_MRI
        assert person["display_name"] == "Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM)"


def _namesakes(identity, monkeypatch, members=()):
    history = [
        {"sender": OTHER_DIR["name"], "sender_mri": OTHER_MRI, "mentions": []},
        {"sender": HOANG_DIR["name"], "sender_mri": HOANG_MRI, "mentions": []},
    ]
    c = _history_client(identity, monkeypatch, history, members)
    monkeypatch.setattr(c, "search_users", _fake_directory([HOANG_DIR, OTHER_DIR]))
    return c


def test_namesakes_with_two_mris_are_listed_not_guessed(identity, monkeypatch):
    from common.errors import Mcp365Error

    c = _namesakes(identity, monkeypatch)
    with pytest.raises(Mcp365Error) as excinfo:
        c.resolve_mentions("19:g@thread.v2", ["Nguyen Huy Hoang"])
    msg = excinfo.value.message
    assert "nhiều người" in msg and "KHÔNG tự chọn" in msg
    assert HOANG_MRI in msg and OTHER_MRI in msg
    assert "(VF-KPTX-VPMDVTM)" in msg and "(VF-KPTX-TKVPTLK)" in msg


def test_namesakes_in_directory_are_listed_with_email(identity, monkeypatch):
    from common.errors import Mcp365Error

    c = _history_client(identity, monkeypatch, [])
    monkeypatch.setattr(c, "search_users", _fake_directory([HOANG_DIR, OTHER_DIR]))
    with pytest.raises(Mcp365Error) as excinfo:
        c.resolve_mentions("19:g@thread.v2", ["Nguyễn Huy Hoàng"])
    assert "hoangnh21@vingroup.net" in excinfo.value.message
    assert "hoangnh99@vingroup.net" in excinfo.value.message


def test_namesakes_narrowed_to_the_one_chat_member(identity, monkeypatch):
    c = _namesakes(identity, monkeypatch, members={HOANG_MRI})
    [person] = c.resolve_mentions("19:g@thread.v2", ["Nguyễn Huy Hoàng"])
    assert person["mri"] == HOANG_MRI


def test_name_with_unit_picks_that_unit(identity, monkeypatch):
    c = _namesakes(identity, monkeypatch)
    [a, b] = c.resolve_mentions(
        "19:g@thread.v2", ["Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM)", "nguyen huy hoang (vf-kptx-tkvptlk)"]
    )
    assert (a["mri"], b["mri"]) == (HOANG_MRI, OTHER_MRI)


def test_name_with_unknown_unit_is_not_matched_loosely(identity, monkeypatch):
    from common.errors import ConversationNotFoundError

    c = _namesakes(identity, monkeypatch)
    with pytest.raises(ConversationNotFoundError):
        c.resolve_mentions("19:g@thread.v2", ["Nguyễn Huy Hoàng (VF-KHAC)"])


@pytest.mark.parametrize(
    "asked", ["hoangnh21@vingroup.net", "v.hoangnh21@vinfast.vn", "hoangnh21", "@hoangnh21", HOANG_MRI,
              "c254a4f9-539c-42e0-94a9-bc6ea1753622", "8:orgid:C254A4F9-539C-42E0-94A9-BC6EA1753622"]
)
def test_mri_email_and_alias_resolve_exactly(identity, monkeypatch, asked):
    c = _namesakes(identity, monkeypatch)
    [person] = c.resolve_mentions("19:g@thread.v2", [asked])
    assert person["mri"] == HOANG_MRI
    assert person["display_name"] == "Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM)"


def test_email_resolves_from_directory_when_not_in_chat(identity, monkeypatch):
    c = _history_client(identity, monkeypatch, [])
    monkeypatch.setattr(c, "search_users", _fake_directory([HOANG_DIR, OTHER_DIR]))
    [person] = c.resolve_mentions("19:g@thread.v2", ["hoangnh21@vingroup.net"])
    assert (person["mri"], person["email"]) == (HOANG_MRI, "hoangnh21@vingroup.net")


@pytest.mark.parametrize(
    ("asked", "hint"),
    [("nobody@vingroup.net", "email"), ("8:orgid:99999999-9999-9999-9999-999999999999", "MRI"), ("Ai Đó", "Ai Đó")],
)
def test_unknown_identity_is_a_clear_error(identity, monkeypatch, asked, hint):
    from common.errors import ConversationNotFoundError

    c = _namesakes(identity, monkeypatch)
    with pytest.raises(ConversationNotFoundError) as excinfo:
        c.resolve_mentions("19:g@thread.v2", [asked])
    assert hint in excinfo.value.message


def test_member_mri_never_seen_asks_for_email(identity, monkeypatch):
    from common.errors import Mcp365Error

    silent = "8:orgid:99999999-9999-9999-9999-999999999999"
    c = _history_client(identity, monkeypatch, [], members={silent})
    with pytest.raises(Mcp365Error) as excinfo:
        c.resolve_mentions("19:g@thread.v2", [silent])
    assert "email" in excinfo.value.remediation


def test_tag_asked_by_alias_lands_on_the_written_name():
    from teams.client import apply_mentions, mention_label

    person = {"name": "hoangnh21", "display_name": "Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM)",
              "mri": HOANG_MRI, "email": "hoangnh21@vingroup.net"}
    html, props = apply_mentions("<p>@Nguyễn Huy Hoàng xem giúp</p>", [person])
    assert html.startswith('<p><span itemtype="http://schema.skype.com/Mention"')
    assert html.endswith(" xem giúp</p>") and "@" not in html
    assert props[0]["mri"] == HOANG_MRI
    assert mention_label(person) == "@Nguyễn Huy Hoàng (VF-KPTX-VPMDVTM) (hoangnh21@vingroup.net)"
    assert mention_label({**person, "email": ""}).endswith("(8:orgid:…a1753622)")


def test_get_members_reads_thread_roster_and_tolerates_refusal(client, monkeypatch):
    from common.errors import Mcp365Error

    monkeypatch.setattr(client, "_chat_json", lambda *a, **k: {"members": [{"id": "8:orgid:ABC"}, {"id": "8:orgid:def"}]})
    assert client.get_members("19:g@thread.v2") == {"8:orgid:abc", "8:orgid:def"}

    def refuse(*a, **k):
        raise Mcp365Error("403", "")

    monkeypatch.setattr(client, "_chat_json", refuse)
    assert client.get_members("19:g@thread.v2") == set()
    assert client.get_members("48:notes") == set()


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


def test_normalize_direct_chat_id():
    from teams.client import normalize_direct_chat_id

    guid_a = "0e5c2252-9da0-4304-9b16-de41ad71ceb2"
    guid_b = "b6cf511d-9f31-4a84-89d8-3a400a1a544f"
    reversed_id = f"19:{guid_b}_{guid_a}@unq.gbl.spaces"
    sorted_id = f"19:{guid_a}_{guid_b}@unq.gbl.spaces"

    assert normalize_direct_chat_id(reversed_id) == sorted_id
    assert normalize_direct_chat_id(sorted_id) == sorted_id
    # Group chats and channels remain unchanged
    assert normalize_direct_chat_id("19:channel123@thread.tacv2") == "19:channel123@thread.tacv2"
    assert normalize_direct_chat_id("19:group123@thread.v2") == "19:group123@thread.v2"


def test_create_or_get_direct_chat(client, monkeypatch):
    auth = client._auth()
    my_guid = auth["identity"].mri.removeprefix("8:orgid:")
    target_guid = "0e5c2252-9da0-4304-9b16-de41ad71ceb2"

    # Self chat returns notes
    assert client.create_or_get_direct_chat(f"8:orgid:{my_guid}") == "48:notes"

    posted = []

    def fake_request(url, method="GET", headers=None, data=None, context="", **kwargs):
        assert "/threads" in url
        assert method == "POST"
        posted.append(url)
        resp_headers = {
            "Location": f"https://apac.ng.msg.teams.microsoft.com/v1/threads/19:{target_guid}_{my_guid}@unq.gbl.spaces"
        }
        return 201, b"{}", resp_headers

    monkeypatch.setattr("teams.client.request", fake_request)
    res = client.create_or_get_direct_chat(f"8:orgid:{target_guid}")
    assert res == f"19:{target_guid}_{my_guid}@unq.gbl.spaces"
    assert len(posted) == 1  # the id came from the Location header, not the local fallback


def test_find_conversation_normalizes_and_resolves_direct_chat(client, monkeypatch):
    guid_a = "0e5c2252-9da0-4304-9b16-de41ad71ceb2"
    guid_b = "b6cf511d-9f31-4a84-89d8-3a400a1a544f"
    reversed_id = f"19:{guid_b}_{guid_a}@unq.gbl.spaces"
    sorted_id = f"19:{guid_a}_{guid_b}@unq.gbl.spaces"

    conv = client.find_conversation(reversed_id)
    assert conv["id"] == sorted_id
    assert conv["type"] == "DirectChat"

    # Directory fallback when user not in recent conversations
    monkeypatch.setattr(client, "list_conversations", lambda **kw: [])
    monkeypatch.setattr(
        client,
        "search_users",
        lambda q, max_results=3: [{"name": "Lê Văn Nguyên", "teams_mri": f"8:orgid:{guid_a}"}],
    )
    monkeypatch.setattr(client, "create_or_get_direct_chat", lambda target: sorted_id)

    conv_by_name = client.find_conversation("Lê Văn Nguyên")
    assert conv_by_name["id"] == sorted_id
    assert "Lê Văn Nguyên" in conv_by_name["name"]


def test_send_message_auto_creates_thread_on_404(client, monkeypatch):
    guid_a = "0e5c2252-9da0-4304-9b16-de41ad71ceb2"
    guid_b = "b6cf511d-9f31-4a84-89d8-3a400a1a544f"
    sorted_id = f"19:{guid_a}_{guid_b}@unq.gbl.spaces"

    monkeypatch.setattr(client, "find_conversation", lambda ident: {"id": sorted_id, "name": "Direct Chat", "type": "DirectChat"})
    monkeypatch.setattr(client, "create_or_get_direct_chat", lambda target: sorted_id)
    monkeypatch.setattr(client, "_resolve_sent_id", lambda *args, **kw: "12345")

    attempts = []
    monkeypatch.setattr(client, "_auth", lambda: {"region": "apac", "token": "t"})

    def fake_request(url, headers=None, method="GET", data=None, context="", **kwargs):
        attempts.append(url)
        if len(attempts) == 1:
            from common.errors import Mcp365Error

            err = Mcp365Error("HTTP 404 Not Found LocationLookupFailed")
            err.http_status = 404
            raise err
        return 201, b'{"OriginalArrivalTime": 1789999999}', {}

    monkeypatch.setattr("teams.client.request", fake_request)

    res = client.send_message(sorted_id, "Xin chào anh")
    assert res["status"] == "SENT"
    assert len(attempts) == 2


# ------------------------------------------------------------ quote replies

# Shape of a real quote reply sent by the Teams client (self chat, 22/09),
# with the author anonymised. The MCP payload must reproduce it exactly.
_AUTHOR = "8:orgid:00000000-0000-0000-0000-00000000000a"
_ORIGINAL = {
    "id": "1785494028836",
    "from": f"https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/{_AUTHOR}",
    "imdisplayname": "Người Gửi (VF-TEST)",
    "originalarrivaltime": "2026-07-31T10:33:48.8360000Z",
    "messagetype": "RichText/Html",
    "content": "<p>WBS-api/docs/task/phase-1</p>",
    "properties": {},
}
_REAL_REPLY_CONTENT = (
    '<blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="1785494028836">\r\n'
    f'<strong itemprop="mri" itemid="{_AUTHOR}">Người Gửi (VF-TEST)</strong>'
    '<span itemprop="time" itemid="1785494028836"></span>\r\n'
    '<p itemprop="preview">WBS-api/docs/task/phase-1</p>\r\n'
    "</blockquote>\r\n"
    "<p>helllo</p>"
)
_REAL_QTD = f'[{{"messageId":"1785494028836","sender":"{_AUTHOR}","time":1785494028836}}]'


def test_reply_quote_matches_what_the_teams_client_sends():
    from teams.client import build_reply_quote

    quote, quoted = build_reply_quote(_ORIGINAL)
    assert quote + "<p>helllo</p>" == _REAL_REPLY_CONTENT
    assert quoted == {"messageId": "1785494028836", "sender": _AUTHOR, "time": 1785494028836}


def test_reply_preview_drops_nested_quotes_and_is_capped():
    from teams.client import quote_preview

    nested = (
        "<p>Cái này đẩy theo đường nào&nbsp;</p>\n"
        '<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="1"><strong itemprop="mri">A</strong>'
        '<p itemprop="preview">code cũ</p></blockquote>'
    )
    assert quote_preview(nested) == "Cái này đẩy theo đường nào"
    long = quote_preview("<p>" + "x" * 300 + "</p>")
    assert len(long) == 200 and long.endswith("…")
    assert quote_preview("<p>a<br>\r\n&nbsp;&nbsp; b &lt;c&gt;</p>") == "a b <c>"


@pytest.fixture
def chat_service(client, monkeypatch):
    """Fake Chat Service: GET of a message returns ``store[id]``; POSTs are captured."""
    import json
    from urllib.parse import unquote

    store = {"1785494028836": dict(_ORIGINAL)}
    posts: list[dict] = []
    monkeypatch.setattr(client, "_auth", lambda: {"region": "apac", "token": "t"})
    monkeypatch.setattr(client, "_resolve_sent_id", lambda *a, **k: "999")

    def fake_request(url, headers=None, method="GET", data=None, context="", **kwargs):
        if method in ("POST", "PUT"):  # sends and edits
            posts.append(json.loads(data))
            return 201, b'{"OriginalArrivalTime": 1790067173993}', {}
        msg_id = unquote(url.rsplit("/", 1)[-1])
        if msg_id in store:
            return 200, json.dumps(store[msg_id]).encode(), {}
        from common.errors import Mcp365Error

        err = Mcp365Error("HTTP 404 Not Found")
        err.http_status = 404
        raise err

    monkeypatch.setattr("teams.client.request", fake_request)
    monkeypatch.setattr(client, "get_messages", lambda *a, **k: {"messages": []})
    return client, store, posts


def test_quote_reply_payload_carries_qtd_msgs(chat_service):
    client, _store, posts = chat_service
    res = client.send_message("48:notes", "helllo", reply_to_id="1785494028836")

    assert res["reply_to_id"] == "1785494028836"
    (payload,) = posts
    assert payload["content"] == _REAL_REPLY_CONTENT
    assert payload["properties"]["qtdMsgs"] == _REAL_QTD
    assert payload["properties"]["formatVariant"] == "TEAMS"
    assert "mentions" not in payload["properties"]


def test_quote_reply_keeps_mentions(chat_service, identity):
    import json

    client, _store, posts = chat_service
    person = {"name": "Hiển", "display_name": "Nguyễn Phú Hiển", "mri": _AUTHOR}
    client.send_message("48:notes", "@Hiển xem giúp", reply_to_id="1785494028836", mentions=[person])

    props = posts[0]["properties"]
    assert json.loads(props["mentions"])[0]["mri"] == _AUTHOR
    assert json.loads(props["qtdMsgs"])[0]["messageId"] == "1785494028836"
    assert posts[0]["content"].startswith('<blockquote itemscope="" itemtype="http://schema.skype.com/Reply"')


def test_quote_reply_to_an_unknown_message_is_not_sent(chat_service):
    from common.errors import Mcp365Error

    client, _store, posts = chat_service
    with pytest.raises(Mcp365Error) as excinfo:
        client.send_message("48:notes", "helllo", reply_to_id="1111111111111")
    assert "CHƯA được gửi" in excinfo.value.message
    assert posts == []  # the old code sent a quote attributed to "Member" instead


def test_quote_reply_to_a_deleted_message_is_not_sent(chat_service):
    from common.errors import Mcp365Error

    client, store, posts = chat_service
    store["1789926857800"] = {"id": "1789926857800", "content": "", "properties": {"deletetime": 1789926865678}}
    with pytest.raises(Mcp365Error):
        client.send_message("48:notes", "helllo", reply_to_id="1789926857800")
    assert posts == []


# ------------------------------------------------------------ edits keep their tags

# Shape of the real group message edited on 23/09 (MRIs anonymised): three tags
# whose @Name the edit must not flatten into plain text.
_DAN, _HANH, _PHUONG = (f"8:orgid:00000000-0000-0000-0000-00000000000{i}" for i in "bcd")
_TAGGED = {
    "id": "1790137061859",
    "content": (
        '<p><span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="0">'
        "Nguyễn Minh Dân (VF-KPTX-VPTAITX)</span>&nbsp;"
        '<span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="1">'
        "Nguyễn Hân Hạnh (VF-KPTX-VPTAITX)</span>&nbsp;"
        '<span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="2">'
        "Nguyễn Việt Phương (VF-KPTX-VPTAITX)</span>&nbsp;em cần chốt mấy điểm</p>"
    ),
    "properties": {
        "mentions": json.dumps(
            [
                {"@type": "http://schema.skype.com/Mention", "itemid": i, "mri": mri, "mentionType": "person",
                 "displayName": f"{name} (VF-KPTX-VPTAITX)"}
                for i, (mri, name) in enumerate(
                    [(_DAN, "Nguyễn Minh Dân"), (_HANH, "Nguyễn Hân Hạnh"), (_PHUONG, "Nguyễn Việt Phương")]
                )
            ],
            ensure_ascii=False,
        )
    },
}


def test_edit_with_mentions_sends_real_tags(chat_service):
    client, _store, posts = chat_service
    person = {"name": "Dân", "display_name": "Nguyễn Minh Dân (VF-KPTX-VPTAITX)", "mri": _DAN}
    res = client.edit_message("48:notes", "1790137061859", "@Dân xem giúp em", mentions=[person])

    (payload,) = posts
    assert payload["content"].startswith(
        '<p><span itemtype="http://schema.skype.com/Mention" itemscope="" itemid="0">'
        "Nguyễn Minh Dân (VF-KPTX-VPTAITX)</span> xem giúp em"
    )
    [tag] = json.loads(payload["properties"]["mentions"])
    assert (tag["mri"], tag["itemid"], tag["displayName"]) == (_DAN, "0", "Nguyễn Minh Dân (VF-KPTX-VPTAITX)")
    assert res["mentioned"] == ["Nguyễn Minh Dân (VF-KPTX-VPTAITX)"]


def test_edit_without_mentions_keeps_the_original_tags_still_written(chat_service):
    """The 23/09 bug: the edit went out as plain "@Nguyễn Minh Dân" text."""
    client, store, posts = chat_service
    store["1790137061859"] = _TAGGED
    new = "@Nguyễn Minh Dân (VF-KPTX-VPTAITX) @Nguyễn Hân Hạnh em cần chốt lại mấy điểm"
    res = client.edit_message("48:notes", "1790137061859", new)

    (payload,) = posts
    tags = json.loads(payload["properties"]["mentions"])
    assert [t["mri"] for t in tags] == [_DAN, _HANH]  # Phương's @Name was removed
    assert "@Nguyễn" not in payload["content"]
    assert "(VF-KPTX-VPTAITX) (VF" not in payload["content"]  # the longest written form became the tag
    assert payload["content"].count("schema.skype.com/Mention") == 2
    assert res["mentioned"] == ["Nguyễn Minh Dân (VF-KPTX-VPTAITX)", "Nguyễn Hân Hạnh (VF-KPTX-VPTAITX)"]


def test_kept_mentions_reports_dropped_tags():
    from teams.client import kept_mentions

    people, dropped = kept_mentions(_TAGGED, "@Nguyễn Việt Phương ơi")
    assert [(p["name"], p["mri"]) for p in people] == [("Nguyễn Việt Phương", _PHUONG)]
    assert dropped == ["Nguyễn Minh Dân (VF-KPTX-VPTAITX)", "Nguyễn Hân Hạnh (VF-KPTX-VPTAITX)"]
    assert kept_mentions({"content": "<p>x</p>", "properties": {}}, "@A") == ([], [])


def test_edit_with_empty_mentions_tags_nobody_and_does_not_read_the_original(chat_service):
    client, _store, posts = chat_service
    client.edit_message("48:notes", "9999999999999", "@Nguyễn Minh Dân thôi", mentions=[])
    (payload,) = posts
    assert "properties" not in payload


def test_edit_whose_original_cannot_be_read_is_not_sent(chat_service):
    from common.errors import Mcp365Error

    client, _store, posts = chat_service
    with pytest.raises(Mcp365Error) as excinfo:
        client.edit_message("48:notes", "1111111111111", "@Ai đó")
    assert "CHƯA được sửa" in excinfo.value.message
    assert posts == []


@pytest.fixture
def edit_tool(chat_service, monkeypatch):
    import anyio
    from mcp.server.mcpserver import MCPServer

    import tools as tools_mod

    client, store, posts = chat_service
    store["1790137061859"] = _TAGGED
    monkeypatch.setattr(client, "search_users", lambda *a, **k: [])
    monkeypatch.setattr(tools_mod, "teams", lambda: client)
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)

    def call(**args):
        return anyio.run(mcp.call_tool, "edit_teams_message", {"chat_name_or_id": "48:notes", **args})

    return call, posts


def test_edit_draft_shows_who_stays_tagged_and_who_is_dropped(edit_tool):
    from mcp.server.mcpserver.exceptions import ToolError

    call, posts = edit_tool
    with pytest.raises(ToolError) as excinfo:
        call(message_id="1790137061859", new_message="@Nguyễn Minh Dân chốt nhé", is_user_confirm=False)
    draft = str(excinfo.value)
    assert "CHƯA GỬI" in draft
    assert "Tag (giữ từ tin gốc):** @Nguyễn Minh Dân (VF-KPTX-VPTAITX)" in draft
    assert "Bỏ tag" in draft and "Nguyễn Việt Phương (VF-KPTX-VPTAITX)" in draft
    assert posts == []


def test_edit_with_unknown_name_fails_clearly_and_edits_nothing(edit_tool):
    from mcp.server.mcpserver.exceptions import ToolError

    call, posts = edit_tool
    with pytest.raises(ToolError) as excinfo:
        call(message_id="1790137061859", new_message="@Người Lạ xem", mentions=["Người Lạ"], is_user_confirm=True)
    assert "Không tìm thấy 'Người Lạ'" in str(excinfo.value)
    assert posts == []


# --------------------------------------- naming a 1:1 chat after the OTHER person

ME_GUID = "b6cf511d-9f31-4a84-89d8-3a400a1a544f"
NAMSON_GUID = "789cece9-3e1d-4d92-b161-1931977e664d"
NAMSON_CHAT = f"19:{NAMSON_GUID}_{ME_GUID}@unq.gbl.spaces"
HIEN_GUID = "e1dcef8e-dae6-468f-8b58-de659169609e"
HIEN_CHAT = f"19:{HIEN_GUID}_{ME_GUID}@unq.gbl.spaces"


def _last(guid: str, name: str) -> dict:
    return {
        "from": f"https://teams.microsoft.com/api/chatsvc/apac/v1/users/ME/contacts/8:orgid:{guid}",
        "imdisplayname": name,
        "content": "hi",
        "messagetype": "RichText/Html",
        "composetime": "2026-09-22T10:21:51.367Z",
    }


def _client(identity, monkeypatch):
    c = TeamsClient()
    monkeypatch.setattr(type(c), "identity", property(lambda self: identity))
    return c


def test_direct_chat_peer_picks_the_other_guid(identity):
    from teams.client import direct_chat_peer

    assert direct_chat_peer(NAMSON_CHAT, identity.mri) == f"8:orgid:{NAMSON_GUID}"
    assert direct_chat_peer(f"19:{ME_GUID}_{HIEN_GUID}@unq.gbl.spaces", identity.mri) == f"8:orgid:{HIEN_GUID}"
    assert direct_chat_peer("19:group@thread.v2", identity.mri) == ""
    assert direct_chat_peer(f"19:{ME_GUID}_{ME_GUID}@unq.gbl.spaces", identity.mri) == ""
    # Neither GUID is me: refuse to guess.
    assert direct_chat_peer(f"19:{NAMSON_GUID}_{HIEN_GUID}@unq.gbl.spaces", identity.mri) == ""


def test_my_own_name_never_labels_a_11_chat(identity, monkeypatch):
    """The real bug: I sent the last message, so the chat took *my* name.

    `1:1 Chat (Nguyễn Phan Nam Sơn)` was rendered as
    `1:1 Chat (Nguyễn Hoàng Sơn (VF-KPTX-VPTAITX))`.
    """
    c = _client(identity, monkeypatch)
    raw = [
        {"id": NAMSON_CHAT, "lastMessage": _last(ME_GUID, identity.display_name)},
        # Nam Sơn spoke last in a group chat on the same page - that is where
        # his name comes from.
        {
            "id": "19:squad@thread.v2",
            "threadProperties": {"topic": "[Vita-S5] Development team"},
            "lastMessage": _last(NAMSON_GUID, "Nguyễn Phan Nam Sơn (VF-KPTX-VPTAITX)"),
        },
    ]
    names = c._learn_peer_names(raw, identity.mri)
    row = c._format_conversation(raw[0], my_mri=identity.mri, peer_names=names)
    assert row["name"] == "1:1 Chat (Nguyễn Phan Nam Sơn (VF-KPTX-VPTAITX))"
    assert identity.display_name not in row["name"]


def test_peer_who_sent_last_is_unchanged(identity, monkeypatch):
    """The chat with anh Hiển always worked; it must keep working."""
    c = _client(identity, monkeypatch)
    conv = {"id": HIEN_CHAT, "lastMessage": _last(HIEN_GUID, "Nguyễn Văn Hiển (VF-KPTX-VPTAITX)")}
    names = c._learn_peer_names([conv], identity.mri)
    row = c._format_conversation(conv, my_mri=identity.mri, peer_names=names)
    assert row["name"] == "1:1 Chat (Nguyễn Văn Hiển (VF-KPTX-VPTAITX))"


def test_unknown_peer_falls_back_to_the_mri_not_to_me(identity, monkeypatch):
    """An honest, useless label beats a confident, wrong one."""
    c = _client(identity, monkeypatch)
    conv = {"id": HIEN_CHAT, "lastMessage": _last(ME_GUID, identity.display_name)}
    row = c._format_conversation(conv, my_mri=identity.mri, peer_names={})
    assert row["name"] == f"1:1 Chat (8:orgid:{HIEN_GUID})"
    assert identity.display_name not in row["name"]


def test_last_sender_still_reports_who_really_spoke(identity, monkeypatch):
    """`last_sender` is data, not a label - the watcher reads it."""
    c = _client(identity, monkeypatch)
    conv = {"id": NAMSON_CHAT, "lastMessage": _last(ME_GUID, identity.display_name)}
    row = c._format_conversation(conv, my_mri=identity.mri, peer_names={})
    assert row["last_sender"] == identity.display_name


def test_peer_names_learned_from_a_chat_read_survive(identity, monkeypatch):
    """Reading a chat teaches the name the listing could not see."""
    c = _client(identity, monkeypatch)
    with c._lock:
        c._peer_names[f"8:orgid:{HIEN_GUID}"] = "Nguyễn Văn Hiển (VF-KPTX-VPTAITX)"
    conv = {"id": HIEN_CHAT, "lastMessage": _last(ME_GUID, identity.display_name)}
    # No explicit map: the client falls back to what it has remembered.
    row = c._format_conversation(conv, my_mri=identity.mri)
    assert row["name"] == "1:1 Chat (Nguyễn Văn Hiển (VF-KPTX-VPTAITX))"


def test_system_events_never_teach_a_peer_name(identity, monkeypatch):
    c = _client(identity, monkeypatch)
    event = {**_last(NAMSON_GUID, "Someone Else"), "messagetype": "ThreadActivity/AddMember"}
    names = c._learn_peer_names([{"id": "19:squad@thread.v2", "lastMessage": event}], identity.mri)
    assert f"8:orgid:{NAMSON_GUID}" not in names


def test_learn_peer_names_never_records_myself(identity, monkeypatch):
    c = _client(identity, monkeypatch)
    names = c._learn_peer_names([{"id": NAMSON_CHAT, "lastMessage": _last(ME_GUID, identity.display_name)}], identity.mri)
    assert identity.mri not in names


def test_self_chat_group_meeting_and_channel_names_are_unchanged(identity, monkeypatch):
    c = _client(identity, monkeypatch)

    notes = c._format_conversation({"id": "48:notes", "lastMessage": _last(ME_GUID, identity.display_name)}, my_mri=identity.mri, peer_names={})
    assert notes["name"] == "Chat with yourself (Notes)"
    assert notes["type"] == "DirectChat"

    group = c._format_conversation(
        {"id": "19:squad@thread.v2", "threadProperties": {"topic": "[Vita-S5] Development team"}},
        my_mri=identity.mri,
        peer_names={},
    )
    assert group["name"] == "[Vita-S5] Development team" and group["type"] == "GroupChat"

    meeting = c._format_conversation(
        {"id": "19:meeting_abc@thread.v2", "threadProperties": {"topic": "Sprint planning"}},
        my_mri=identity.mri,
        peer_names={},
    )
    assert meeting["name"] == "Sprint planning" and meeting["type"] == "MeetingChat"

    channel = c._format_conversation(
        {
            "id": "19:chan@thread.tacv2",
            "threadProperties": {"spaceThreadTopic": "Vita-S5", "topicThreadTopic": "General"},
        },
        my_mri=identity.mri,
        peer_names={},
    )
    assert channel["name"] == "[Vita-S5] #General" and channel["type"] == "Channel"

    # System feeds are still dropped.
    assert c._format_conversation({"id": "48:calllogs"}, my_mri=identity.mri, peer_names={}) is None


def test_bot_chat_keeps_the_last_sender_label(identity, monkeypatch):
    """A `28:` bot has no GUID pair, so there is nothing better than its name."""
    c = _client(identity, monkeypatch)
    conv = {
        "id": "19:bot_thread@unq.gbl.spaces",
        "lastMessage": {"from": ".../contacts/28:app-id", "imdisplayname": "Workflows", "content": ""},
    }
    row = c._format_conversation(conv, my_mri=identity.mri, peer_names={})
    assert row["name"] == "1:1 Chat (Workflows)"


# ------------------------------------------------ 1:1 messages in the briefing

def _conv(conv_id: str, name: str, kind: str) -> dict:
    return {"id": conv_id, "name": name, "type": kind, "last_activity": "", "last_sender": "", "last_message": ""}


def _feed_msg(mri: str, name: str, hours_ago: float, text: str) -> dict:
    from datetime import UTC

    return {"sender": name, "sender_mri": mri, "timestamp_dt": datetime.now(UTC) - timedelta(hours=hours_ago), "content": text}


@pytest.fixture
def dm_feed(identity, monkeypatch):
    c = _client(identity, monkeypatch)
    old_chat = f"19:{ME_GUID}_{'0' * 8}-0000-0000-0000-{'0' * 12}@unq.gbl.spaces"
    convs = [
        _conv("48:notes", "Chat with yourself (Notes)", "DirectChat"),
        _conv(HIEN_CHAT, "1:1 Chat (Hiền)", "DirectChat"),
        _conv(old_chat, "1:1 Chat (Cũ)", "DirectChat"),
        _conv(NAMSON_CHAT, "1:1 Chat (Nam Sơn)", "DirectChat"),
        _conv("19:abc@thread.v2", "Dev team", "GroupChat"),
    ]
    me = f"8:orgid:{ME_GUID.upper()}"  # as sent: case and prefix are not guaranteed
    history = {
        "48:notes": [_feed_msg(me, "Sơn", 1, "ghi chú")],
        HIEN_CHAT: [_feed_msg(me, "Sơn", 1, "em gửi chị rồi nhé")],
        old_chat: [_feed_msg(f"8:orgid:{'0' * 8}-0000-0000-0000-{'0' * 12}", "Cũ", 30, "tuần trước")],
        NAMSON_CHAT: [
            _feed_msg(f"8:orgid:{NAMSON_GUID}", "Nam Sơn", 2, "anh xem giúp em PR"),
            _feed_msg(me, "Sơn", 1, "ok để anh xem"),
        ],
        "19:abc@thread.v2": [_feed_msg(me, "Sơn", 1, "chốt nhé")],
    }
    fetched: list[str] = []
    monkeypatch.setattr(c, "list_conversations", lambda **kw: list(convs))

    def fake_get(conv_id, limit):
        fetched.append(conv_id)
        return {"messages": history[conv_id]}

    monkeypatch.setattr(c, "get_messages", fake_get)
    return c, fetched


def test_direct_feed_lists_only_chats_someone_else_wrote_in(dm_feed):
    c, fetched = dm_feed
    res = c.get_recent_feed(hours=24, max_chats=3, chat_types=("DirectChat",), incoming_only=True)
    # The notes chat is dropped before truncating, so it costs no slot and no request.
    assert "48:notes" not in fetched and NAMSON_CHAT in fetched
    # Hiền: only I spoke. Cũ: the peer spoke before the cutoff.
    assert [f["chat_name"] for f in res["feed"]] == ["1:1 Chat (Nam Sơn)"]
    # My reply stays as context.
    assert [m["content"] for m in res["feed"][0]["messages"]] == ["anh xem giúp em PR", "ok để anh xem"]


def test_group_feed_is_unchanged_by_default(dm_feed):
    c, _fetched = dm_feed
    res = c.get_recent_feed(hours=24)
    assert [f["chat_name"] for f in res["feed"]] == ["Dev team"]


@pytest.mark.anyio
async def test_daily_briefing_shows_untagged_direct_messages(monkeypatch):
    from mcp.server.mcpserver import MCPServer

    import tools as tools_mod

    feed_calls: list[dict] = []

    class FakeTeams:
        def get_user_mentions(self, **kw):
            return {"mentions": [], "errors": []}

        def get_recent_feed(self, **kw):
            feed_calls.append(kw)
            if kw.get("chat_types") == ("DirectChat",):
                item = {"chat_name": "1:1 Chat (Nam Sơn)", "messages": [{"sender": "Nam Sơn", "content": "anh xem giúp em PR"}]}
                return {"feed": [item], "errors": [], "scanned": 1}
            return {"feed": [], "errors": [], "scanned": 0}

        def get_calendar_events(self, start, end):
            return []

    class FakeSharePoint:
        def search_files(self, **kw):
            return []

    monkeypatch.setattr(tools_mod, "teams", lambda: FakeTeams())
    monkeypatch.setattr(tools_mod, "sp", lambda: FakeSharePoint())
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)
    res = await mcp.call_tool("get_daily_briefing", {"hours": 24})
    text = res.content[0].text
    assert "## 📨 2. Tin nhắn 1:1 (1 cuộc trò chuyện)" in text
    assert "### 👤 **1:1 Chat (Nam Sơn)**" in text and "anh xem giúp em PR" in text
    assert "## 💬 3. Thảo luận tại các nhóm" in text and "## 📄 5." in text
    assert any(kw.get("incoming_only") for kw in feed_calls)
