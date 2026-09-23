"""Block until new Outlook mail arrives in the Inbox, then print it and exit.

The mail counterpart of ``watch.py``: run in the background by an agent
harness, it exits the moment a matching mail lands, and that exit is what wakes
the agent up.

    bin/mcp-365-mail-watch                                     # any new mail in the Inbox
    bin/mcp-365-mail-watch --from "nam son" --subject "review" --timeout 1500
    bin/mcp-365-mail-watch --unread-only --important --since 2026-09-22T02:00:00Z

Read-only: this never marks mail as read, moves, flags, deletes or sends it.

Re-arming: the last line of the output is the exact ``--since ... --seen-ids ...``
to run next. ``--since`` is set :data:`REARM_GRACE` *before* the newest mail
shown, and ``--seen-ids`` lists the mails already shown inside that window, so a
second mail in the same second, or one delivered late with an older
``ReceivedDateTime``, is still caught - and nothing is shown twice.

Every request goes through ``OutlookMailClient.list_messages`` (the code path of
the ``list_emails`` tool), authenticated from the Chrome session like the MCP
server. At most ``--scan`` (<= 50) newest mails are checked per poll.

Exit codes: 0 new mail printed · 3 nothing within --timeout · 1 auth/config error.
"""

from __future__ import annotations

import argparse
import logging
import sys
import time
from collections.abc import Sequence
from datetime import UTC, datetime, timedelta, timezone
from typing import Any, NamedTuple

from common.errors import AuthExpiredError, ConfigError, CookieError, Mcp365Error
from outlook.client import OutlookMailClient
from teams.client import _parse_timestamp, fold

EXIT_FOUND, EXIT_ERROR, EXIT_TIMEOUT = 0, 1, 3
VN = timezone(timedelta(hours=7), "ICT")  # Asia/Ho_Chi_Minh, no DST
MAX_SCAN = 50  # OutlookMailClient.list_messages caps $top at 50 and does not paginate.
#: How far behind the newest mail seen the next cursor starts. Exchange stamps
#: ``ReceivedDateTime`` on delivery, but a mail can become visible to the REST
#: listing a little after a newer one; the window re-reads that stretch and the
#: seen-id list keeps it from being shown twice.
REARM_GRACE = timedelta(minutes=2)

#: Needs the user to act (sign in again, fix config); retrying cannot help.
FATAL = (AuthExpiredError, ConfigError, CookieError)
AUTH_HINT = (
    "Mở https://outlook.office.com trong Chrome (đúng profile đã cấu hình), đăng nhập tài khoản "
    "công ty và chọn 'Stay signed in', rồi chạy lại watcher. Kiểm tra nhanh bằng tool "
    "`check_365_connection`."
)


def received_at(msg: dict[str, Any]) -> datetime | None:
    return _parse_timestamp(msg.get("received") or "")


def relevant_messages(
    messages: list[dict[str, Any]],
    since: datetime,
    *,
    from_names: Sequence[str] = (),
    subjects: Sequence[str] = (),
    unread_only: bool = False,
    important: bool = False,
    flagged: bool = False,
) -> list[dict[str, Any]]:
    """Mails received at or after ``since`` that pass every given filter, oldest first.

    ``>=`` matches the server's ``ReceivedDateTime ge``. Already-shown mails are
    *not* removed here - that is :func:`poll_once`'s job, by id - because the
    re-arm cursor deliberately overlaps the previous window (:data:`REARM_GRACE`):
    a same-second or late-delivered mail is only safe from loss if the cursor
    does not jump past it. Each filter narrows; repeating ``--from`` or
    ``--subject`` widens that filter (any of them).
    """
    wanted_from = [fold(n) for n in from_names if n.strip()]
    wanted_subject = [fold(s) for s in subjects if s.strip()]
    out = []
    for msg in messages:
        ts = received_at(msg)
        if ts is None or ts < since:
            continue
        if wanted_from and not any(w in fold(msg.get("from", "")) for w in wanted_from):
            continue
        if wanted_subject and not any(w in fold(msg.get("subject", "")) for w in wanted_subject):
            continue
        if unread_only and msg.get("is_read"):
            continue
        if important and str(msg.get("importance", "")).casefold() != "high":
            continue
        if flagged and not msg.get("is_flagged"):
            continue
        out.append(msg)
    out.sort(key=received_at)
    return out


def _utc_text(ts: datetime) -> str:
    # Floors to the whole second, which only widens the window; ids dedupe it.
    return ts.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%SZ")


def rearm(matched: list[dict[str, Any]], since: datetime) -> tuple[str, list[str]]:
    """The ``(--since, --seen-ids)`` to re-arm with after this poll.

    ``matched`` is every mail of the poll that passed the filters, shown now or
    earlier. The cursor goes back :data:`REARM_GRACE` from the newest of them,
    never before the current ``since`` (this run has not looked further back),
    and every matched mail inside the new window is listed as seen - the list
    is complete because this poll saw that whole window.
    """
    stamps = [received_at(m) for m in matched]
    newest = max((ts for ts in stamps if ts), default=since)
    cursor = _utc_text(max(since, newest - REARM_GRACE))
    floor = _parse_timestamp(cursor)
    seen = [str(m.get("id")) for m, ts in zip(matched, stamps, strict=True) if ts and ts >= floor and m.get("id")]
    return cursor, seen


def rearm_command(cursor: str, seen: Sequence[str]) -> str:
    ids = ",".join(seen)
    return f"--since {cursor}" + (f" --seen-ids '{ids}'" if ids else "")


def next_since(found: list[dict[str, Any]], since: datetime | None = None) -> str:
    """The ``--since`` part of :func:`rearm` (kept for callers that only need the cursor)."""
    return rearm(found, since or datetime.min.replace(tzinfo=UTC))[0]


def render(
    found: list[dict[str, Any]],
    *,
    truncated: bool = False,
    scan: int = MAX_SCAN,
    rearm_args: str = "",
) -> str:
    lines = [f"📧 {len(found)} mail mới"]
    for msg in found:
        when = received_at(msg).astimezone(VN).strftime("%d/%m %H:%M")
        tags = []
        if not msg.get("is_read"):
            tags.append("📩 chưa đọc")
        if str(msg.get("importance", "")).casefold() == "high":
            tags.append("❗ quan trọng")
        if msg.get("is_flagged"):
            tags.append("🚩")
        if msg.get("has_attachments"):
            tags.append("📎")
        tag = f" [{' · '.join(tags)}]" if tags else ""
        lines.append(f"- [{when}] {msg.get('from') or '?'} · {msg.get('subject') or '(không có tiêu đề)'}{tag}")
        preview = " ".join((msg.get("preview") or "").split())
        if preview:
            lines.append(f"  > {preview[:160]}{'…' if len(preview) > 160 else ''}")
        lines.append(f"  id `{msg.get('id')}`")
    if truncated:
        lines.append(
            f"⚠️ Đã quét đủ {scan} mail mới nhất; mail cũ hơn trong khoảng này có thể chưa được xét. "
            "Dùng `list_emails` nếu cần xem hết."
        )
    rearm_args = rearm_args or rearm_command(*rearm(found, min(received_at(m) for m in found)))
    lines.append(f"→ Đọc: `read_email(message_id)`. Theo dõi tiếp: `{rearm_args}`")
    return "\n".join(lines)


class Poll(NamedTuple):
    found: list[dict[str, Any]]  # matching and not shown before: what wakes the agent
    matched: list[dict[str, Any]]  # every mail passing the filters, shown before or not
    newest: datetime | None  # newest ReceivedDateTime on the page, matching or not
    truncated: bool  # the page came back full: older mails in the window were not read


def poll_once(
    client: Any, since: datetime, args: argparse.Namespace, seen: frozenset[str] = frozenset()
) -> Poll:
    """One Inbox listing: the matching mails, minus those already shown (by id)."""
    # Whole seconds for the server filter: it is only a lower bound, the exact
    # comparison happens in relevant_messages.
    messages = client.list_messages(
        folder="inbox",
        limit=args.scan,
        unread_only=args.unread_only,
        since=since.replace(microsecond=0).isoformat(),
    )
    matched = relevant_messages(
        messages,
        since,
        from_names=args.from_names,
        subjects=args.subjects,
        unread_only=args.unread_only,
        important=args.important,
        flagged=args.flagged,
    )
    found = [m for m in matched if str(m.get("id")) not in seen]
    stamps = [ts for ts in (received_at(m) for m in messages) if ts]
    return Poll(found, matched, max(stamps, default=None), len(messages) >= args.scan)


def _parse_seen(values: Sequence[str]) -> frozenset[str]:
    ids = (part.strip().strip("'\"") for value in values for part in value.split(","))
    return frozenset(i for i in ids if i)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("--from", dest="from_names", action="append", default=[],
                   help="Only wake for this sender's name or email (repeatable, diacritics optional).")
    p.add_argument("--subject", dest="subjects", action="append", default=[],
                   help="Only wake when the subject contains this keyword (repeatable, diacritics optional).")
    p.add_argument("--unread-only", action="store_true", help="Ignore mail already read.")
    p.add_argument("--important", action="store_true", help="Only mail marked High importance.")
    p.add_argument("--flagged", action="store_true", help="Only flagged mail.")
    p.add_argument("--interval", type=float, default=60, help="Seconds between polls (default 60).")
    p.add_argument("--timeout", type=float, default=1500, help="Give up after this many seconds (default 1500).")
    p.add_argument("--since", default="", help="ISO time to watch from, UTC unless it has an offset (default: now).")
    p.add_argument("--seen-ids", dest="seen_ids", action="append", default=[],
                   help="Comma-separated mail ids already shown; copy it from the re-arm line (repeatable).")
    p.add_argument("--scan", type=int, default=MAX_SCAN, help=f"Newest mails checked per poll (1-{MAX_SCAN}).")
    args = p.parse_args(argv)
    args.scan = max(1, min(args.scan, MAX_SCAN))
    return args


def main(argv: list[str] | None = None, *, sleep=time.sleep) -> int:
    args = parse_args(argv)
    logging.basicConfig(
        level=logging.WARNING, format="%(asctime)s %(name)s: %(message)s", datefmt="%H:%M:%S", stream=sys.stderr
    )
    since = _parse_timestamp(args.since) if args.since else datetime.now(UTC)
    if since is None:
        print(f"--since không hợp lệ: {args.since}", file=sys.stderr)
        return EXIT_ERROR

    seen = _parse_seen(args.seen_ids)
    last_matched: list[dict[str, Any]] | None = None
    client = OutlookMailClient()
    deadline = time.monotonic() + args.timeout
    while True:
        try:
            poll = poll_once(client, since, args, seen)
        except FATAL as exc:
            # Printed to stdout on purpose: the agent must be woken to tell the user.
            print(f"⚠️ Watcher mail dừng vì lỗi xác thực/cấu hình: {exc}\nCách sửa chung: {AUTH_HINT}")
            return EXIT_ERROR
        except Mcp365Error as exc:
            print(f"(bỏ qua lỗi tạm thời: {exc.message})", file=sys.stderr)
            poll = None
        if poll and poll.found:
            args_next = rearm_command(*rearm(poll.matched, since))
            print(render(poll.found, truncated=poll.truncated, scan=args.scan, rearm_args=args_next), flush=True)
            return EXIT_FOUND
        if poll:
            last_matched = poll.matched
            # Nothing new. With a complete page, move the cursor up behind the
            # newest mail examined (keeping the grace window), so a busy Inbox
            # cannot fill the page with mail this run already ruled out. A full
            # page means older mails in the window were never read: keep it.
            if poll.newest and not poll.truncated:
                since = max(since, poll.newest - REARM_GRACE)
        if time.monotonic() + args.interval > deadline:
            local = since.astimezone(VN)
            if last_matched is None:
                args_next = rearm_command(_utc_text(since), sorted(seen))
            else:
                args_next = rearm_command(*rearm(last_matched, since))
            print(
                f"⏱️ Không có mail mới trong {int(args.timeout)}s (theo dõi từ {local:%d/%m %H:%M} giờ VN). "
                f"Theo dõi tiếp: `{args_next}`"
            )
            return EXIT_TIMEOUT
        sleep(args.interval)


if __name__ == "__main__":
    sys.exit(main())
