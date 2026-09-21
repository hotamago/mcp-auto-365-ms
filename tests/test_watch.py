"""The background watcher's filtering, verified without a network."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta

import watch

SINCE = datetime(2026, 9, 21, 5, 0, tzinfo=UTC)
ME = "8:orgid:me"


def _msg(minutes: int, sender: str, mri: str, mentions_me: bool = False, text: str = "hi") -> dict:
    return {
        "id": f"m{minutes}",
        "sender": sender,
        "sender_mri": mri,
        "timestamp_dt": SINCE + timedelta(minutes=minutes),
        "content": text,
        "mentions_me": mentions_me,
    }


GROUP = {"id": "19:dev@thread.v2", "name": "[Vita-S5] Development team", "type": "GroupChat"}
DM = {"id": "19:a_b@unq.gbl.spaces", "name": "1:1 Chat (Nam Sơn)", "type": "DirectChat"}


def _relevant(conv, messages, **kw):
    opts = {"watched_ids": set(), "want_dm": False, "want_mentions": False, "from_names": []}
    opts.update(kw)
    return watch.relevant_messages(conv, messages, SINCE, ME, **opts)


def test_only_conversations_active_after_the_cursor_are_fetched():
    convs = [
        {"id": "old", "last_activity": "2026-09-21T04:59:00Z"},
        {"id": "new", "last_activity": "2026-09-21T05:01:00Z"},
        {"id": "none", "last_activity": ""},
    ]
    assert [c["id"] for c in watch.active_since(convs, SINCE)] == ["new"]


def test_own_messages_and_older_ones_never_wake_the_agent():
    msgs = [_msg(-1, "Hùng", "8:orgid:hung"), _msg(2, "Me", ME), _msg(3, "Hùng", "8:orgid:hung")]
    assert [m["id"] for m in _relevant(GROUP, msgs, watched_ids={GROUP["id"]})] == ["m3"]


def test_watched_chat_can_be_narrowed_to_one_sender_without_diacritics():
    msgs = [_msg(1, "Trịnh Anh Tuấn", "8:orgid:tuan"), _msg(2, "Phạm Sỹ Hùng (VF-KPTX)", "8:orgid:hung")]
    got = _relevant(GROUP, msgs, watched_ids={GROUP["id"]}, from_names=["pham sy hung"])
    assert [m["id"] for m in got] == ["m2"]


def test_busy_unwatched_group_only_wakes_on_a_mention():
    msgs = [_msg(1, "Tuấn", "8:orgid:tuan"), _msg(2, "Tuấn", "8:orgid:tuan", mentions_me=True)]
    assert [m["id"] for m in _relevant(GROUP, msgs, want_mentions=True)] == ["m2"]


def test_any_new_direct_message_wakes_with_dm_flag():
    msgs = [_msg(1, "Nam Sơn", "8:orgid:ns")]
    assert _relevant(DM, msgs, want_dm=True) == msgs
    assert _relevant(DM, msgs) == []


def test_defaults_to_dm_and_mentions_when_nothing_is_specified():
    args = watch.parse_args([])
    assert args.dm and args.mentions
    args = watch.parse_args(["--chat", "Dev"])
    assert not args.dm and not args.mentions


def test_render_shows_chat_sender_and_ids():
    out = watch.render([(GROUP, _msg(3, "Phạm Sỹ Hùng", "8:orgid:hung", text="Vậy chạy  L40s\nnhé em"))])
    assert out.startswith("🔔 1 tin mới")
    assert "[Vita-S5] Development team · Phạm Sỹ Hùng: Vậy chạy L40s nhé em" in out
    assert "id `m3`" in out
