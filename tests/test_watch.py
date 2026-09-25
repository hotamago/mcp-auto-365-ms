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


# ------------------------------------------------------------ độ nóng hội thoại


def _event(seconds: float, weight: float = 1.0):
    return (SINCE + timedelta(seconds=seconds), weight)


def _weight(conv, msg, **kw):
    opts = {"watched_ids": set(), "want_dm": False, "want_mentions": False, "from_names": []}
    opts.update(kw)
    return watch.message_weight(conv, msg, ME, **opts)


def test_heat_halves_every_half_life_and_ignores_clock_skew():
    events = [_event(0)]
    assert watch.heat(events, SINCE, 300) == 1.0
    assert watch.heat(events, SINCE + timedelta(seconds=300), 300) == 0.5
    assert watch.heat(events, SINCE + timedelta(seconds=900), 300) == 0.125
    # Một tin "ở tương lai" (đồng hồ lệch) không được nặng hơn tin vừa gửi.
    assert watch.heat([_event(30)], SINCE, 300) == 1.0


def test_heat_adds_up_every_message():
    assert watch.heat([_event(0), _event(0), _event(0, 0.5)], SINCE, 300) == 2.5
    two_old = [_event(-300), _event(-300)]
    assert watch.heat(two_old, SINCE, 300) == watch.heat([_event(0)], SINCE, 300)
    assert watch.heat([], SINCE, 300) == 0.0


def test_poll_interval_steps_down_to_the_floor_and_back_to_the_ceiling():
    table = {0: 60, 0.49: 60, 0.5: 30, 1.4: 30, 2: 20, 3: 15, 4: 12, 5: 10, 50: 10}
    for h, expected in table.items():
        assert watch.poll_interval(h, slow=60, fastest=10) == expected, h
    assert watch.poll_interval(1, slow=30, fastest=10) == 15
    assert watch.poll_interval(9, slow=30, fastest=10) == 10
    # Sàn không thấp hơn trần: không bao giờ tăng tốc, và không bao giờ chậm hơn --interval.
    assert watch.poll_interval(9, slow=60, fastest=60) == 60
    assert watch.poll_interval(9, slow=20, fastest=60) == 20


def test_direct_messages_and_mentions_weigh_more_than_group_talk():
    notes = {"id": "48:notes", "name": "Notes", "type": "DirectChat"}
    peer, mine = _msg(1, "Nam Sơn", "8:orgid:ns"), _msg(2, "Me", ME)
    assert _weight(DM, peer, want_dm=True) == _weight(DM, mine, want_dm=True) == watch.W_DIRECT
    assert _weight(DM, peer, watched_ids={DM["id"]}) == watch.W_DIRECT
    tagged = _msg(3, "Tuấn", "8:orgid:tuan", mentions_me=True)
    assert _weight(GROUP, tagged, want_mentions=True) == watch.W_MENTION
    assert _weight(GROUP, mine, want_mentions=True) == watch.W_GROUP < watch.W_MENTION
    hung, tuan = _msg(4, "Phạm Sỹ Hùng", "8:orgid:hung"), _msg(5, "Tuấn", "8:orgid:tuan")
    watched = {"watched_ids": {GROUP["id"]}, "from_names": ["pham sy hung"]}
    assert _weight(GROUP, hung, **watched) == watch.W_GROUP
    assert _weight(GROUP, tuan, **watched) == 0.0
    # Nhóm ồn ào mà mình không được tag, và ghi chú cho chính mình, không làm nóng gì cả.
    assert _weight(GROUP, tuan, want_mentions=True, want_dm=True) == 0.0
    assert _weight(notes, _msg(6, "Me", ME), want_dm=True) == 0.0


def test_own_messages_heat_the_chat_but_never_wake_the_agent():
    mine = [_msg(1, "Me", ME), _msg(2, "Me", ME)]
    assert _relevant(DM, mine, want_dm=True) == []
    events = watch.heat_events(DM, mine, ME, watched_ids=set(), want_dm=True, want_mentions=False, from_names=[])
    assert [w for _, w in events] == [watch.W_DIRECT, watch.W_DIRECT]


def test_a_chat_is_reread_only_when_its_last_activity_moves():
    state = watch.Heat()
    conv = {**DM, "last_activity": "2026-09-21T05:02:00Z"}
    msgs = [_msg(1, "Nam Sơn", "8:orgid:ns"), _msg(2, "Me", ME)]
    state.record(conv, msgs, [])
    assert state.is_fresh(conv)
    assert not state.is_fresh({**conv, "last_activity": "2026-09-21T05:03:00Z"})


def test_a_listing_ahead_of_the_history_is_reread_once_more():
    state = watch.Heat()
    conv = {**DM, "last_activity": "2026-09-21T05:05:00Z"}  # tin 05:05 chưa hiện trong lịch sử
    state.record(conv, [_msg(1, "Nam Sơn", "8:orgid:ns")], [])
    assert not state.is_fresh(conv)
    state.record(conv, [_msg(1, "Nam Sơn", "8:orgid:ns")], [])  # vẫn chưa có: thôi, coi như tin hệ thống
    assert state.is_fresh(conv)


def test_pace_note_names_the_hot_chat_and_says_when_it_cooled():
    at = SINCE
    assert '"1:1 Chat (Nam Sơn)" nóng lên (độ nóng 2.1)' in watch.pace_note(60, 20, 2.1, "1:1 Chat (Nam Sơn)", 60, at)
    assert "poll mỗi 30s: nguội bớt (độ nóng 1.2)" in watch.pace_note(20, 30, 1.2, "x", 60, at)
    assert "poll mỗi 60s: đã nguội" in watch.pace_note(30, 60, 0.4, "x", 60, at)


# ------------------------------------------------------- vòng lặp với đồng hồ giả


class _Clock:
    """Đồng hồ giả: ``sleep`` đẩy cả đồng hồ đơn điệu lẫn giờ thật lên, không ngủ thật."""

    def __init__(self):
        self.t = 0.0
        self.sleeps: list[float] = []

    def monotonic(self) -> float:
        return self.t

    def now(self) -> datetime:
        return SINCE + timedelta(seconds=self.t)

    def sleep(self, seconds: float) -> None:
        self.sleeps.append(seconds)
        self.t += seconds


def _at(seconds: float, sender: str, mri: str, **kw) -> dict:
    msg = _msg(0, sender, mri, **kw)
    msg["id"] = f"s{seconds:g}"
    msg["timestamp_dt"] = SINCE + timedelta(seconds=seconds)
    return msg


class _Teams:
    """Thay TeamsClient: lịch sử chat hiện dần theo đồng hồ giả; đếm lần liệt kê và lần đọc tin."""

    def __init__(self, clock: _Clock, chats: list[tuple[dict, list[dict]]]):
        self.clock = clock
        self.chats = chats
        self.identity = type("Identity", (), {"mri": ME})()
        self.listings = 0
        self.fetches: list[tuple[float, str]] = []

    def _visible(self, conv_id: str) -> list[dict]:
        history = next(msgs for conv, msgs in self.chats if conv["id"] == conv_id)
        return [m for m in history if m["timestamp_dt"] <= self.clock.now()]

    def find_conversation(self, ident: str) -> dict:
        return next(conv for conv, _ in self.chats if ident in (conv["id"], conv["name"]))

    def list_conversations(self, page_size: int = 50, use_cache: bool = True) -> list[dict]:
        self.listings += 1
        out = []
        for conv, _ in self.chats:
            seen = self._visible(conv["id"])
            last = max((m["timestamp_dt"] for m in seen), default=None)
            out.append({**conv, "last_activity": last.isoformat().replace("+00:00", "Z") if last else ""})
        return out

    def get_messages(self, conv_id: str, limit: int = 30) -> dict:
        self.fetches.append((self.clock.t, conv_id))
        return {"messages": self._visible(conv_id)[-limit:]}


def _run(monkeypatch, chats, argv):
    clock = _Clock()
    teams = _Teams(clock, chats)
    monkeypatch.setattr(watch, "TeamsClient", lambda: teams)
    code = watch.main(
        ["--since", SINCE.isoformat(), *argv], sleep=clock.sleep, clock=clock.monotonic, now=clock.now
    )
    return code, clock, teams


#: Vừa trao đổi xong ngay trước --since: agent vừa được đánh thức và bật lại watcher.
RECENT_DM = [
    _at(-120, "Nam Sơn", "8:orgid:ns"),
    _at(-60, "Me", ME),
    _at(-30, "Nam Sơn", "8:orgid:ns"),
]


def test_restart_after_a_hot_exchange_polls_fast_then_relaxes(monkeypatch, capsys):
    code, clock, teams = _run(monkeypatch, [(DM, RECENT_DM)], ["--dm", "--timeout", "3600"])
    err = capsys.readouterr().err
    assert code == watch.EXIT_TIMEOUT
    # Làm ấm đọc chat 1:1 đó một lần (tin cũ hơn --since nên không đánh thức), rồi không đọc lại.
    assert teams.fetches == [(0.0, DM["id"])]
    assert teams.listings == len(clock.sleeps) + 1
    assert clock.sleeps[0] == 15  # độ nóng ~2.6 → 60 / (1 + 3)
    assert clock.sleeps == sorted(clock.sleeps) and clock.sleeps[-1] == 60
    assert sum(clock.sleeps) <= 3600
    # Chỉ ghi log khi đổi bậc: 15 → 20 → 30 → 60, không phải mỗi vòng.
    lines = [line for line in err.splitlines() if "poll mỗi" in line]
    assert len(lines) == 4
    assert '"1:1 Chat (Nam Sơn)" nóng lên' in lines[0] and "đã nguội" in lines[-1]


def test_no_adaptive_polls_exactly_like_before(monkeypatch, capsys):
    code, clock, teams = _run(monkeypatch, [(DM, RECENT_DM)], ["--dm", "--timeout", "600", "--no-adaptive"])
    assert code == watch.EXIT_TIMEOUT
    assert set(clock.sleeps) == {60} and teams.fetches == []
    assert "poll mỗi" not in capsys.readouterr().err


def test_cold_chats_keep_the_old_pace_and_print_nothing_extra(monkeypatch, capsys):
    old = [_at(-3600, "Nam Sơn", "8:orgid:ns")]
    code, clock, teams = _run(monkeypatch, [(DM, old)], ["--dm", "--timeout", "300"])
    captured = capsys.readouterr()
    assert code == watch.EXIT_TIMEOUT
    assert set(clock.sleeps) == {60} and teams.fetches == []
    assert "poll mỗi" not in captured.err
    assert captured.out.startswith("⏱️ Không có tin mới trong 300s")


def test_my_own_reply_speeds_up_polling_and_the_answer_wakes_the_agent(monkeypatch, capsys):
    history = [_at(30, "Me", ME), _at(100, "Nam Sơn", "8:orgid:ns", text="ok để mình xem")]
    code, clock, teams = _run(monkeypatch, [(DM, history)], ["--dm", "--timeout", "1500"])
    out = capsys.readouterr().out
    assert code == watch.EXIT_FOUND
    # 60 s nguội; thấy tin mình lúc 60 s → 30 s; 90 s chat không đổi nên không đọc lại; 120 s có tin đáp.
    assert clock.sleeps == [60, 30, 30]
    assert teams.fetches == [(60.0, DM["id"]), (120.0, DM["id"])]
    assert "ok để mình xem" in out and "id `s30`" not in out


def test_timeout_is_kept_while_the_interval_changes(monkeypatch, capsys):
    code, clock, _ = _run(monkeypatch, [(DM, RECENT_DM)], ["--dm", "--timeout", "100"])
    assert code == watch.EXIT_TIMEOUT
    # Nhịp 15 → 20 s: dừng khi lần ngủ tiếp theo (theo đúng nhịp lúc đó) sẽ vượt hạn chót.
    assert clock.sleeps == [15, 20, 20, 20, 20]
    assert "Không có tin mới trong 100s" in capsys.readouterr().out
    _, clock, _ = _run(monkeypatch, [(DM, RECENT_DM)], ["--dm", "--timeout", "100", "--no-adaptive"])
    assert clock.sleeps == [60]


def test_busy_group_without_a_tag_neither_heats_nor_gets_reread(monkeypatch):
    chatter = [_at(s, "Tuấn", "8:orgid:tuan") for s in (-50, -40, 10)]
    code, clock, teams = _run(monkeypatch, [(GROUP, chatter)], ["--mentions", "--timeout", "300"])
    assert code == watch.EXIT_TIMEOUT
    assert set(clock.sleeps) == {60}
    # Vòng đầu đọc để làm ấm; vòng 60 s đọc tin 10 s; sau đó last_activity không đổi nên thôi đọc.
    assert teams.fetches == [(0.0, GROUP["id"]), (60.0, GROUP["id"])]


def test_warmup_reads_only_the_most_recent_relevant_chats(monkeypatch):
    other = {"id": "19:c_d@unq.gbl.spaces", "name": "1:1 Chat (Hiển)", "type": "DirectChat"}
    third = {"id": "19:e_f@unq.gbl.spaces", "name": "1:1 Chat (Nguyệt)", "type": "DirectChat"}
    chats = [
        (DM, [_at(-60, "Nam Sơn", "8:orgid:ns")]),
        (other, [_at(-30, "Hiển", "8:orgid:hien")]),
        (third, [_at(-2000, "Nguyệt", "8:orgid:nguyet")]),  # ngoài 3 half-life: không đọc
        (GROUP, [_at(-10, "Tuấn", "8:orgid:tuan")]),  # không theo dõi: không đọc
    ]
    _, _, teams = _run(monkeypatch, chats, ["--dm", "--timeout", "0", "--warmup", "1"])
    assert teams.fetches == [(0.0, other["id"])]
    # Với --mentions nhóm cũng liên quan, nhưng chat 1:1 được làm ấm trước dù cũ hơn.
    _, _, teams = _run(monkeypatch, chats, ["--dm", "--mentions", "--timeout", "0", "--warmup", "2"])
    assert teams.fetches == [(0.0, other["id"]), (0.0, DM["id"])]


def test_a_failed_poll_records_no_heat_so_nothing_is_skipped_later():
    class _Flaky:
        identity = type("Identity", (), {"mri": ME})()

        def list_conversations(self, **_kw):
            return [{**DM, "last_activity": "2026-09-21T05:01:00Z"}, {**GROUP, "last_activity": "2026-09-21T05:01:00Z"}]

        def get_messages(self, conv_id, limit=30):
            if conv_id == GROUP["id"]:
                raise watch.Mcp365Error("HTTP 503", "Thử lại.")
            return {"messages": [_msg(1, "Me", ME)]}

    state = watch.Heat()
    args = watch.parse_args(["--dm", "--mentions"])
    try:
        watch.poll_once(_Flaky(), SINCE, args, set(), state)
    except watch.Mcp365Error:
        pass
    assert state.fetched == {} and state.events == {}


def test_new_options_have_sensible_defaults():
    args = watch.parse_args([])
    assert (args.interval, args.min_interval, args.half_life, args.warmup, args.adaptive) == (60, 10, 300, 8, True)
    assert watch.parse_args(["--no-adaptive"]).adaptive is False
