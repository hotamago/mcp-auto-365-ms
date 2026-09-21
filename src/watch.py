"""Block until new Teams messages worth reacting to arrive, then print them and exit.

Meant to be run in the background by an agent harness: the process exits the
moment something relevant lands, and that exit is what wakes the agent up. So
the agent can wait for a reply without the user having to prompt it.

    bin/mcp-365-watch                                  # 1:1 messages + mentions of me
    bin/mcp-365-watch --chat "[Vita-S5] Development team" --from "Phạm Sỹ Hùng"
    bin/mcp-365-watch --dm --mentions --timeout 1500

Read-only: this never sends, reacts or edits anything.

Exit codes: 0 new messages printed · 3 nothing within --timeout · 1 auth/config error.
"""

from __future__ import annotations

import argparse
import sys
import time
from datetime import UTC, datetime
from typing import Any

from common.errors import AuthExpiredError, ConfigError, Mcp365Error
from common.identity import normalize_mri
from teams.client import TeamsClient, _parse_timestamp, fold

EXIT_FOUND, EXIT_ERROR, EXIT_TIMEOUT = 0, 1, 3


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


def relevant_messages(
    conv: dict[str, Any],
    messages: list[dict[str, Any]],
    since: datetime,
    me_mri: str,
    *,
    watched_ids: set[str],
    want_dm: bool,
    want_mentions: bool,
    from_names: list[str],
) -> list[dict[str, Any]]:
    """Messages in ``conv`` newer than ``since`` that the user would want to hear about."""
    me = normalize_mri(me_mri)
    watched = conv["id"] in watched_ids
    is_dm = conv.get("type") == "DirectChat" and conv["id"] != "48:notes"
    wanted_from = [fold(n) for n in from_names]

    out = []
    for msg in messages:
        ts = msg.get("timestamp_dt")
        if ts is None or ts <= since:
            continue
        if me and normalize_mri(msg.get("sender_mri", "")) == me:
            continue  # our own messages never wake us up
        sender_ok = not wanted_from or any(w in fold(msg.get("sender", "")) for w in wanted_from)
        if (watched and sender_ok) or (want_dm and is_dm) or (want_mentions and msg.get("mentions_me")):
            out.append(msg)
    return out


def render(found: list[tuple[dict[str, Any], dict[str, Any]]]) -> str:
    lines = [f"🔔 {len(found)} tin mới"]
    for conv, msg in found:
        when = msg["timestamp_dt"].astimezone().strftime("%H:%M")
        text = " ".join((msg.get("content") or "").split())
        files = ", ".join(f["name"] for f in msg.get("attachments") or [])
        if files:
            text = f"{text} 📎 {files}".strip()
        tag = " 🔔mention" if msg.get("mentions_me") else ""
        lines.append(
            f"- [{when}] {conv['name']} · {msg.get('sender', '?')}{tag}: {text[:400]} "
            f"(chat `{conv['id']}` · id `{msg.get('id')}`)"
        )
    return "\n".join(lines)


def poll_once(client: TeamsClient, since: datetime, args: argparse.Namespace, watched_ids: set[str]):
    conversations = client.list_conversations(page_size=0, use_cache=False)
    found = []
    for conv in active_since(conversations, since):
        is_dm = conv.get("type") == "DirectChat"
        if not (conv["id"] in watched_ids or (args.dm and is_dm) or args.mentions):
            continue
        messages = client.get_messages(conv["id"], limit=args.scan)["messages"]
        for msg in relevant_messages(
            conv,
            messages,
            since,
            client.identity.mri,
            watched_ids=watched_ids,
            want_dm=args.dm,
            want_mentions=args.mentions,
            from_names=args.from_names,
        ):
            found.append((conv, msg))
    found.sort(key=lambda pair: pair[1]["timestamp_dt"])
    return found


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__.split("\n\n")[0])
    p.add_argument("--chat", action="append", default=[], help="Chat name or id to watch (repeatable).")
    p.add_argument("--from", dest="from_names", action="append", default=[],
                   help="Only wake for these senders in --chat chats (repeatable, diacritics optional).")
    p.add_argument("--dm", action="store_true", help="Wake on any new 1:1 message.")
    p.add_argument("--mentions", action="store_true", help="Wake on any message that mentions me.")
    p.add_argument("--interval", type=float, default=60, help="Seconds between polls (default 60).")
    p.add_argument("--timeout", type=float, default=1500, help="Give up after this many seconds (default 1500).")
    p.add_argument("--since", default="", help="ISO time to watch from (default: now).")
    p.add_argument("--scan", type=int, default=30, help="Messages fetched per active chat (default 30).")
    args = p.parse_args(argv)
    if not (args.chat or args.dm or args.mentions):
        args.dm = args.mentions = True
    return args


def main(argv: list[str] | None = None) -> int:
    args = parse_args(argv)
    since = _parse_timestamp(args.since) if args.since else datetime.now(UTC)
    if since is None:
        print(f"--since không hợp lệ: {args.since}", file=sys.stderr)
        return EXIT_ERROR

    client = TeamsClient()
    try:
        watched_ids = {client.find_conversation(c)["id"] for c in args.chat}
    except Mcp365Error as exc:
        print(f"⚠️ Watcher không khởi động được: {exc}")
        return EXIT_ERROR

    deadline = time.monotonic() + args.timeout
    while True:
        try:
            found = poll_once(client, since, args, watched_ids)
        except (AuthExpiredError, ConfigError) as exc:
            # Printed to stdout on purpose: the agent must be woken to tell the user.
            print(f"⚠️ Watcher dừng vì lỗi xác thực: {exc}")
            return EXIT_ERROR
        except Mcp365Error as exc:
            print(f"(bỏ qua lỗi tạm thời: {exc.message})", file=sys.stderr)
            found = []
        if found:
            print(render(found), flush=True)
            return EXIT_FOUND
        if time.monotonic() + args.interval > deadline:
            print(f"⏱️ Không có tin mới trong {int(args.timeout)}s (theo dõi từ {since:%H:%M} UTC).")
            return EXIT_TIMEOUT
        time.sleep(args.interval)


if __name__ == "__main__":
    sys.exit(main())
