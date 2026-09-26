"""Block until new Teams messages worth reacting to arrive, then print them and exit.

Meant to be run in the background by an agent harness: the process exits the
moment something relevant lands, and that exit is what wakes the agent up. So
the agent can wait for a reply without the user having to prompt it.

    bin/mcp-365-watch                                  # 1:1 messages + mentions of me
    bin/mcp-365-watch --chat "[Vita-S5] Development team" --from "Phạm Sỹ Hùng"
    bin/mcp-365-watch --dm --mentions --timeout 1500
    bin/mcp-365-watch --dm --mentions --replies-to-me --digest-chat "S5 Development team" --settle 10

Read-only: this never sends, reacts or edits anything.

Hai mức tin (26/09 - bớt số lần đánh thức agent):

- **Mức 1, đánh thức:** tin 1:1 (``--dm``), tin tag mình (``--mentions``), tin trả lời trích dẫn
  một tin của mình (``--replies-to-me``), và mọi tin ở ``--chat`` (lọc ``--from``). Sau tin mức 1
  đầu tiên chờ thêm ``--settle`` giây (mặc định 0 = thoát ngay như cũ) để gom tin tới liền sau.
- **Mức 2, gom:** mọi tin khác của người khác trong ``--digest-chat``. Không thoát ngay; thoát khi
  tin mức 2 chờ lâu nhất đã nằm trong watcher quá ``--digest-after`` giây (tính từ lúc watcher
  thấy tin, không theo giờ gửi: danh sách chat của Teams hay trễ vài phút). Có tin mức 1 thì in
  kèm mọi tin mức 2 đang chờ.
- In ra: dòng ``🔔 N tin mới`` (N = tổng), rồi phần ``🔔 cần xem`` và ``💬 tin nhóm``. Hết
  ``--timeout`` mà còn tin đang chờ thì vẫn in ra (mã 0), không bỏ tin nào.

Every request goes through ``TeamsClient``, so the watcher shares the MCP
server's Chat Service endpoint router (fastest endpoint first, failover,
cooldown). Endpoint choices and switches are logged to stderr; stdout stays
reserved for the wake-up message.

Nhịp poll thích nghi theo độ nóng hội thoại (tắt bằng ``--no-adaptive``):

- Mỗi vòng vẫn chỉ 1 lần liệt kê hội thoại; chỉ đọc tin của chat liên quan có
  ``last_activity`` mới hơn ``--since`` *và* khác lần đọc trước trong lần chạy này.
- Độ nóng một chat = tổng trọng số các tin gần đây × 0.5^(tuổi / ``--half-life``):
  tin 1:1 (cả tin của mình) và tin tag mình 1.0; tin mình gửi trong nhóm và tin ở
  ``--chat`` lọt ``--from`` 0.5; còn lại 0. Tin của mình chỉ làm nóng, không đánh thức.
- Nhịp = ``--interval`` / (1 + độ nóng tổng làm tròn), không dưới ``--min-interval``:
  mặc định 60 → 30 → 20 → 15 → 12 → 10 s. Nguội thì tự về ``--interval``.
- Không có file trạng thái: độ nóng dựng lại từ lịch sử chat. Vòng đầu "làm ấm" bằng
  cách đọc thêm tối đa ``--warmup`` chat liên quan có hoạt động trong 3 half-life gần
  nhất, nên chạy lại ngay sau khi được đánh thức vẫn giữ được nhịp nhanh.
- Mỗi lần đổi nhịp ghi một dòng ngắn ra stderr.

Exit codes: 0 new messages printed · 3 nothing within --timeout · 1 auth/config error.
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from datetime import UTC, datetime, timedelta
from typing import Any

from common.errors import AuthExpiredError, ConfigError, Mcp365Error
from common.identity import normalize_mri
from teams.client import TeamsClient, _parse_timestamp, fold

EXIT_FOUND, EXIT_ERROR, EXIT_TIMEOUT = 0, 1, 3

#: Trọng số một tin khi tính độ nóng, trước khi suy giảm theo tuổi.
W_DIRECT = 1.0  # tin trong chat 1:1 đang theo dõi, cả tin của mình: người kia hay đáp lại ngay
W_MENTION = 1.0  # tin tag mình
W_GROUP = 0.5  # tin mình gửi trong nhóm, hoặc tin ở --chat lọt bộ lọc --from
#: Làm ấm chỉ đọc chat có hoạt động trong chừng này half-life: tin cũ hơn còn ≤ 1/8 trọng số.
WARM_HALF_LIVES = 3

Event = tuple[datetime, float]


def active_since(conversations: list[dict[str, Any]], since: datetime) -> list[dict[str, Any]]:
    """Conversations whose last activity is newer than ``since``.

    One conversation listing tells us where anything happened, so only those
    chats are fetched - polling every chat every minute would be ~30 requests.
    """
    out = []
    for conv in conversations:
        ts = _parse_timestamp(conv.get("last_activity", ""))
        if ts and ts > since:
            out.append(conv)
    return out


def _is_dm(conv: dict[str, Any]) -> bool:
    return conv.get("type") == "DirectChat" and conv["id"] != "48:notes"


def _sender_ok(msg: dict[str, Any], wanted_from: list[str]) -> bool:
    return not wanted_from or any(w in fold(msg.get("sender", "")) for w in wanted_from)


def my_message_ids(messages: list[dict[str, Any]], me_mri: str) -> frozenset[str]:
    """Id các tin của chính mình trong lịch sử vừa đọc (để nhận ra trích dẫn chỉ ghi id)."""
    me = normalize_mri(me_mri)
    return frozenset(str(m.get("id")) for m in messages if me and normalize_mri(m.get("sender_mri", "")) == me)


def quotes_me(msg: dict[str, Any], me_mri: str, my_ids: frozenset[str] = frozenset()) -> bool:
    """Tin này trả lời trích dẫn một tin của mình.

    Tác giả tin được trích lấy từ ``quotes`` (``TeamsClient.parse_quotes``: ``qtdMsgs.sender``,
    hoặc ``<strong itemprop="mri">`` trong blockquote). Trích dẫn không ghi tác giả thì chỉ tính
    khi id tin được trích là một tin của mình trong lịch sử vừa đọc - không đoán theo tên.
    """
    me = normalize_mri(me_mri)
    if not me:
        return False
    for quote in msg.get("quotes") or []:
        author = normalize_mri(quote.get("sender_mri", ""))
        if author:
            if author == me:
                return True
        elif quote.get("message_id") and str(quote["message_id"]) in my_ids:
            return True
    return False


def classify_messages(
    conv: dict[str, Any],
    messages: list[dict[str, Any]],
    since: datetime,
    me_mri: str,
    *,
    watched_ids: set[str],
    want_dm: bool,
    want_mentions: bool,
    from_names: list[str],
    want_replies: bool = False,
    digest_ids: set[str] = frozenset(),
) -> list[tuple[int, dict[str, Any]]]:
    """``(mức, tin)`` cho các tin của người khác trong ``conv`` mới hơn ``since``: 1 đánh thức, 2 gom."""
    me = normalize_mri(me_mri)
    watched = conv["id"] in watched_ids
    digest = conv["id"] in digest_ids
    is_dm = _is_dm(conv)
    wanted_from = [fold(n) for n in from_names]
    mine = my_message_ids(messages, me_mri) if want_replies else frozenset()

    out = []
    for msg in messages:
        ts = msg.get("timestamp_dt")
        if ts is None or ts <= since:
            continue
        if me and normalize_mri(msg.get("sender_mri", "")) == me:
            continue  # our own messages never wake us up
        if (
            (watched and _sender_ok(msg, wanted_from))
            or (want_dm and is_dm)
            or (want_mentions and msg.get("mentions_me"))
            or (want_replies and quotes_me(msg, me_mri, mine))
        ):
            out.append((1, msg))
        elif digest:
            out.append((2, msg))
    return out


def relevant_messages(conv: dict[str, Any], messages: list[dict[str, Any]], since: datetime, me_mri: str,
                      **mode: Any) -> list[dict[str, Any]]:
    """Messages in ``conv`` newer than ``since`` that wake the agent (level 1)."""
    return [msg for level, msg in classify_messages(conv, messages, since, me_mri, **mode) if level == 1]


# ------------------------------------------------------------ độ nóng (hàm thuần)


def message_weight(
    conv: dict[str, Any],
    msg: dict[str, Any],
    me_mri: str,
    *,
    watched_ids: set[str],
    want_dm: bool,
    want_mentions: bool,
    from_names: list[str],
    want_replies: bool = False,
    digest_ids: set[str] = frozenset(),
    my_ids: frozenset[str] = frozenset(),
) -> float:
    """Một tin làm nóng hội thoại bao nhiêu (0 = không). Không liên quan tới việc đánh thức.

    Tin của mình được tính: mình vừa nói thì dễ sắp có người đáp. Nhưng nó vẫn không
    đánh thức agent - việc đó chỉ ``relevant_messages`` quyết. Trong nhóm không theo dõi
    (chỉ đọc vì ``--mentions``) hay nhóm gom (``--digest-chat``), tin người khác không tag
    hay trả lời mình không làm nóng: một nhóm ồn ào không được kéo nhịp poll lên.
    """
    if conv["id"] == "48:notes":
        return 0.0  # ghi chú cho chính mình: không ai trả lời
    me = normalize_mri(me_mri)
    mine = bool(me) and normalize_mri(msg.get("sender_mri", "")) == me
    watched = conv["id"] in watched_ids
    if _is_dm(conv) and (want_dm or watched):
        return W_DIRECT
    if mine:
        return W_GROUP
    if want_mentions and msg.get("mentions_me"):
        return W_MENTION
    if want_replies and quotes_me(msg, me_mri, my_ids):
        return W_MENTION
    if watched and _sender_ok(msg, [fold(n) for n in from_names]):
        return W_GROUP
    return 0.0


def heat_events(conv: dict[str, Any], messages: list[dict[str, Any]], me_mri: str, **mode: Any) -> list[Event]:
    """``(thời điểm, trọng số)`` của các tin làm nóng ``conv``, cả tin cũ hơn ``--since``."""
    out = []
    my_ids = my_message_ids(messages, me_mri) if mode.get("want_replies") else frozenset()
    for msg in messages:
        ts = msg.get("timestamp_dt")
        weight = message_weight(conv, msg, me_mri, my_ids=my_ids, **mode)
        if ts is not None and weight > 0:
            out.append((ts, weight))
    return out


def heat(events: Iterable[Event], now: datetime, half_life: float) -> float:
    """Σ trọng số × 0.5^(tuổi / half_life). Tin "ở tương lai" (lệch đồng hồ) tính tuổi 0."""
    total = 0.0
    for ts, weight in events:
        age = max(0.0, (now - ts).total_seconds())
        total += weight * 0.5 ** (age / half_life)
    return total


def poll_interval(heat_value: float, *, slow: float, fastest: float) -> float:
    """Nhịp poll cho độ nóng: ``slow / (1 + n)``, n = độ nóng làm tròn, không dưới ``fastest``.

    Nhịp theo bậc (mặc định 60/30/20/15/12/10 s) chứ không liên tục, nên chỉ đổi - và chỉ
    ghi log - khi độ nóng qua một mốc. Độ nóng < 0.5 là nguội: về ``slow``.
    ``fastest >= slow`` nghĩa là không bao giờ tăng tốc.
    """
    level = int(heat_value + 0.5)
    return max(min(fastest, slow), slow / (1 + level))


@dataclass
class Heat:
    """Độ nóng các hội thoại trong một lần chạy - chỉ trong bộ nhớ, dựng lại từ lịch sử chat.

    ``events`` giữ tin đã đọc của từng chat nên chat không có gì mới vẫn nguội dần theo
    đồng hồ mà không phải đọc lại. ``fetched`` nhớ ``last_activity`` đã đọc đủ: chat không
    đổi ``last_activity`` thì không đọc lại, nên poll dày hơn không kéo theo đọc nhiều hơn.
    """

    events: dict[str, list[Event]] = field(default_factory=dict)
    names: dict[str, str] = field(default_factory=dict)
    fetched: dict[str, tuple[str, bool]] = field(default_factory=dict)

    def is_fresh(self, conv: dict[str, Any]) -> bool:
        """Đã đọc chat này ở đúng ``last_activity`` hiện tại: không có tin mới."""
        return self.fetched.get(conv["id"]) == (conv.get("last_activity", ""), True)

    def record(self, conv: dict[str, Any], messages: list[dict[str, Any]], events: list[Event]) -> None:
        cid = conv["id"]
        self.events[cid] = events
        self.names[cid] = conv.get("name") or cid
        stamp = conv.get("last_activity", "")
        last = _parse_timestamp(stamp)
        newest = max((m["timestamp_dt"] for m in messages if m.get("timestamp_dt")), default=None)
        seen_newest = last is None or (newest is not None and newest >= last)
        # Không thấy tin mang đúng last_activity (tin hệ thống bị lọc, hoặc danh sách đi trước
        # lịch sử một nhịp): đọc lại thêm một lần rồi mới coi là đã đọc đủ.
        retried = self.fetched.get(cid, ("", False))[0] == stamp
        self.fetched[cid] = (stamp, seen_newest or retried)

    def total(self, now: datetime, half_life: float) -> tuple[float, str]:
        """Độ nóng tổng (cộng mọi chat) và tên chat nóng nhất."""
        per_chat = {cid: heat(ev, now, half_life) for cid, ev in self.events.items()}
        if not per_chat:
            return 0.0, ""
        hottest = max(per_chat, key=per_chat.__getitem__)
        return sum(per_chat.values()), self.names.get(hottest, hottest)


def pace_note(old: float, new: float, heat_value: float, hottest: str, slow: float, when: datetime) -> str:
    """Một dòng stderr khi nhịp poll đổi bậc."""
    clock = when.astimezone().strftime("%H:%M")
    if new >= slow:
        why = "đã nguội"
    elif new < old:
        why = f'"{hottest}" nóng lên (độ nóng {heat_value:.1f})'
    else:
        why = f"nguội bớt (độ nóng {heat_value:.1f})"
    return f"({clock} poll mỗi {new:g}s: {why})"


# ------------------------------------------------------------------ poll


Found = list[tuple[dict[str, Any], dict[str, Any]]]


def _line(conv: dict[str, Any], msg: dict[str, Any], me_mri: str = "") -> str:
    when = msg["timestamp_dt"].astimezone().strftime("%H:%M")
    text = " ".join((msg.get("content") or "").split())
    files = ", ".join(f["name"] for f in msg.get("attachments") or [])
    if files:
        text = f"{text} 📎 {files}".strip()
    tag = " 🔔mention" if msg.get("mentions_me") else ""
    if me_mri and quotes_me(msg, me_mri):
        tag += " ↩️trả lời mình"
    return (
        f"- [{when}] {conv['name']} · {msg.get('sender', '?')}{tag}: {text[:400]} "
        f"(chat `{conv['id']}` · id `{msg.get('id')}`)"
    )


def render(found: Found, digest: Found = (), me_mri: str = "") -> str:
    """Dòng đầu ``🔔 N tin mới`` với N = tổng mọi tin (resume_watcher đếm theo nó), rồi hai phần."""
    found, digest = list(found), list(digest)
    head = f"🔔 {len(found) + len(digest)} tin mới"
    if not digest:
        return "\n".join([head] + [_line(c, m, me_mri) for c, m in found])
    parts = ([f"cần xem {len(found)}"] if found else []) + [f"tin nhóm {len(digest)}"]
    lines = [f"{head} ({' · '.join(parts)})"]
    if found:
        lines.append(f"🔔 cần xem ({len(found)})")
        lines += [_line(c, m, me_mri) for c, m in found]
    lines.append(f"💬 tin nhóm ({len(digest)})")
    lines += [_line(c, m, me_mri) for c, m in digest]
    return "\n".join(lines)


def _watching(conv: dict[str, Any], args: argparse.Namespace, watched_ids: set[str],
              digest_ids: set[str] = frozenset()) -> bool:
    return (
        conv["id"] in watched_ids
        or conv["id"] in digest_ids
        or (args.dm and conv.get("type") == "DirectChat")
        or args.mentions
        or args.replies_to_me
    )


def poll_once(
    client: TeamsClient,
    since: datetime,
    args: argparse.Namespace,
    watched_ids: set[str],
    heat_state: Heat | None = None,
    warm_from: datetime | None = None,
    digest_ids: set[str] = frozenset(),
) -> list[tuple[int, dict[str, Any], dict[str, Any]]]:
    """Một vòng: 1 lần liệt kê hội thoại, rồi đọc tin của chat liên quan có hoạt động mới.

    Trả về ``(mức, chat, tin)`` theo thời gian: mức 1 đánh thức, mức 2 gom (``--digest-chat``).

    Có ``heat_state``: chat đã đọc ở đúng ``last_activity`` hiện tại thì bỏ qua, và mọi tin
    đọc được (cả tin cũ, cả tin của mình) cập nhật độ nóng - chỉ sau khi cả vòng thành công,
    để lỗi giữa chừng không làm vòng sau bỏ sót chat. ``warm_from`` (vòng đầu): đọc thêm tối
    đa ``args.warmup`` chat liên quan hoạt động sau mốc đó; tin của chúng không mới hơn
    ``since`` nên chỉ dùng cho độ nóng, không đánh thức.
    """
    conversations = client.list_conversations(page_size=0, use_cache=False)
    targets = [c for c in active_since(conversations, since) if _watching(c, args, watched_ids, digest_ids)]
    if warm_from is not None and warm_from < since:
        taken = {c["id"] for c in targets}
        extra = [
            c
            for c in active_since(conversations, warm_from)
            if c["id"] not in taken and _watching(c, args, watched_ids, digest_ids)
        ]
        # Chat 1:1 và --chat trước, nhóm chỉ đọc vì --mentions sau (tin người khác ở đó thường
        # nặng 0), rồi mới tới độ mới: làm ấm lặp lại sau mỗi lần được đánh thức, nên phải đáng tiền.
        extra.sort(key=lambda c: (c["id"] in watched_ids or _is_dm(c), _parse_timestamp(c["last_activity"])),
                   reverse=True)
        targets += extra[: max(0, args.warmup)]
    if heat_state is not None:
        targets = [c for c in targets if not heat_state.is_fresh(c)]

    mode = {
        "watched_ids": watched_ids,
        "want_dm": args.dm,
        "want_mentions": args.mentions,
        "from_names": args.from_names,
        "want_replies": args.replies_to_me,
        "digest_ids": digest_ids,
    }
    found, read = [], []
    for conv in targets:
        messages = client.get_messages(conv["id"], limit=args.scan)["messages"]
        read.append((conv, messages))
        for level, msg in classify_messages(conv, messages, since, client.identity.mri, **mode):
            found.append((level, conv, msg))
    if heat_state is not None:
        for conv, messages in read:
            heat_state.record(conv, messages, heat_events(conv, messages, client.identity.mri, **mode))
    found.sort(key=lambda item: item[2]["timestamp_dt"])
    return found


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("--chat", action="append", default=[], help="Chat name or id to watch (repeatable).")
    p.add_argument("--from", dest="from_names", action="append", default=[],
                   help="Only wake for these senders in --chat chats (repeatable, diacritics optional).")
    p.add_argument("--dm", action="store_true", help="Wake on any new 1:1 message.")
    p.add_argument("--mentions", action="store_true", help="Wake on any message that mentions me.")
    p.add_argument("--replies-to-me", action="store_true",
                   help="Đánh thức khi có tin trả lời trích dẫn một tin của mình (mọi chat).")
    p.add_argument("--digest-chat", dest="digest_chats", action="append", default=[],
                   help="Nhóm gom (lặp được): tin thường ở đây không đánh thức ngay, gom rồi in sau --digest-after.")
    p.add_argument("--settle", type=float, default=0,
                   help="Sau tin mức 1 đầu tiên, chờ thêm chừng này giây gom tin tới liền sau (mặc định 0 = thoát ngay).")
    p.add_argument("--digest-after", type=float, default=60,
                   help="Thoát khi tin nhóm gom chờ lâu nhất đã nằm trong watcher quá chừng này giây (mặc định 60).")
    p.add_argument("--interval", type=float, default=60,
                   help="Nhịp poll khi không chat nào nóng, cũng là nhịp chậm nhất (giây, mặc định 60).")
    p.add_argument("--min-interval", type=float, default=10,
                   help="Nhịp nhanh nhất khi hội thoại đang nóng - sàn chống bị Teams throttle (giây, mặc định 10).")
    p.add_argument("--half-life", type=float, default=300,
                   help="Sau chừng này giây, trọng số một tin trong độ nóng còn một nửa (mặc định 300).")
    p.add_argument("--warmup", type=int, default=8,
                   help="Số chat liên quan hoạt động gần đây đọc thêm ở vòng đầu để có độ nóng (mặc định 8, 0 = bỏ).")
    p.add_argument("--no-adaptive", dest="adaptive", action="store_false",
                   help="Poll đều mỗi --interval giây như cũ: không tính độ nóng, không làm ấm.")
    p.add_argument("--timeout", type=float, default=1500, help="Give up after this many seconds (default 1500).")
    p.add_argument("--since", default="", help="ISO time to watch from (default: now).")
    p.add_argument("--scan", type=int, default=30, help="Messages fetched per active chat (default 30).")
    args = p.parse_args(argv)
    if not (args.chat or args.dm or args.mentions or args.replies_to_me or args.digest_chats):
        args.dm = args.mentions = True
    if args.half_life <= 0:
        p.error("--half-life phải lớn hơn 0")
    if args.settle < 0 or args.digest_after < 0:
        p.error("--settle và --digest-after không được âm")
    return args


def main(
    argv: list[str] | None = None,
    *,
    sleep: Callable[[float], None] = time.sleep,
    clock: Callable[[], float] = time.monotonic,
    now: Callable[[], datetime] = lambda: datetime.now(UTC),
) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.WARNING, format="%(asctime)s %(name)s: %(message)s", datefmt="%H:%M:%S", stream=sys.stderr
    )
    logging.getLogger("teams.endpoints").setLevel(logging.INFO)
    since = _parse_timestamp(args.since) if args.since else now()
    if since is None:
        print(f"--since không hợp lệ: {args.since}", file=sys.stderr)
        return EXIT_ERROR

    client = TeamsClient()
    try:
        watched_ids = {client.find_conversation(c)["id"] for c in args.chat}
        digest_ids = {client.find_conversation(c)["id"] for c in args.digest_chats} - watched_ids
    except Mcp365Error as exc:
        print(f"⚠️ Watcher không khởi động được: {exc}")
        return EXIT_ERROR

    heat_state = Heat() if args.adaptive else None
    warm = heat_state is not None and args.warmup > 0
    interval = args.interval
    deadline = clock() + args.timeout
    #: Tin đang chờ in: (chat id, tin id) -> (mức, chat, tin, lúc watcher thấy tin lần đầu).
    pending: dict[tuple[str, str], tuple[int, dict[str, Any], dict[str, Any], float]] = {}
    settle_at: float | None = None

    def flush() -> int:
        items = sorted(pending.values(), key=lambda item: item[2]["timestamp_dt"])
        wake = [(conv, msg) for level, conv, msg, _ in items if level == 1]
        digest = [(conv, msg) for level, conv, msg, _ in items if level == 2]
        print(render(wake, digest, client.identity.mri), flush=True)
        return EXIT_FOUND

    while True:
        try:
            warm_from = now() - timedelta(seconds=WARM_HALF_LIVES * args.half_life) if warm else None
            found = poll_once(client, since, args, watched_ids, heat_state, warm_from, digest_ids)
            warm = False  # làm ấm một lần, sau vòng đầu thành công
        except (AuthExpiredError, ConfigError) as exc:
            # Printed to stdout on purpose: the agent must be woken to tell the user.
            # Tin đang chờ không in: mốc của resume_watcher giữ nguyên nên lần sau báo lại đủ.
            print(f"⚠️ Watcher dừng vì lỗi xác thực: {exc}")
            return EXIT_ERROR
        except Mcp365Error as exc:
            print(f"(bỏ qua lỗi tạm thời: {exc.message})", file=sys.stderr)
            found = []
        seen_at = clock()
        for level, conv, msg in found:
            key = (conv["id"], str(msg.get("id")))
            if key not in pending or level < pending[key][0]:
                pending[key] = (level, conv, msg, pending[key][3] if key in pending else seen_at)
            if level == 1 and settle_at is None:
                settle_at = seen_at + args.settle
        oldest = min((item[3] for item in pending.values() if item[0] == 2), default=None)
        digest_at = oldest + args.digest_after if oldest is not None else None
        due = [at for at in (settle_at, digest_at) if at is not None]
        if due and seen_at >= min(due):
            return flush()
        if heat_state is not None:
            at = now()
            level, hottest = heat_state.total(at, args.half_life)
            pace = poll_interval(level, slow=args.interval, fastest=args.min_interval)
            if pace != interval:
                print(pace_note(interval, pace, level, hottest, args.interval, at), file=sys.stderr, flush=True)
                interval = pace
        # Có tin đang chờ thì dậy đúng hạn gom; nhịp đổi theo độ nóng, nên hạn chót so với đúng lần ngủ sắp tới.
        wait = min([interval] + [max(0.0, at - seen_at) for at in due])
        if clock() + wait > deadline:
            if pending:
                return flush()  # hết giờ vẫn in tin đang chờ, không bỏ tin nào
            print(f"⏱️ Không có tin mới trong {int(args.timeout)}s (theo dõi từ {since:%H:%M} UTC).")
            return EXIT_TIMEOUT
        sleep(wait)


if __name__ == "__main__":
    sys.exit(main())
