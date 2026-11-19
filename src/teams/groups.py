"""Teams group roster, topic and personal conversation properties. No message sends."""

from __future__ import annotations

import json
import time
from typing import TYPE_CHECKING
from urllib.parse import quote, unquote, urlsplit

from common.errors import Mcp365Error
from common.http import decode_json
from common.multi import run_ordered

from .client import _ms, _sender_mri, bot_mri, direct_chat_peer

if TYPE_CHECKING:
    from .client import TeamsClient


def group_path(chat: str, *, meeting: bool = False) -> str:
    """``/threads/<id>`` of a normal group chat; ``meeting`` also lets a meeting chat through (leaving only)."""
    if not chat.startswith("19:") or not chat.endswith("@thread.v2") or (chat.startswith("19:meeting_") and not meeting):
        raise Mcp365Error("Thao tác này chỉ dành cho nhóm chat thường" + (" hoặc chat cuộc họp." if meeting else "."),
                          "Dùng list_teams_chats để chọn đúng nhóm.")
    return f"/threads/{quote(chat, safe='')}"


def _wire(body: dict) -> dict:
    return {"data": json.dumps(body, ensure_ascii=False).encode(), "json_body": True}


def _member_status(result: dict | list, mris: list[str]) -> str:
    """Do not treat a missing/unknown member status as success."""
    rows = result.get("members", result) if isinstance(result, dict) else result
    if isinstance(rows, list):
        rows = {str(r.get("id") or r.get("mri") or r.get("memberId")).lower(): r for r in rows if isinstance(r, dict)}
    elif isinstance(rows, dict):
        rows = {str(k).lower(): v for k, v in rows.items()}
    lines = []
    for mri in mris:
        row = rows.get(mri.lower()) if isinstance(rows, dict) else None
        code = row.get("statusCode", row.get("status", row.get("code"))) if isinstance(row, dict) else row
        label = ("đã thêm" if str(code).isdigit() and 200 <= int(code) < 300 else
                 "bị từ chối" if str(code).isdigit() else "chưa xác nhận")
        lines.append(f"{mri}: {label}" + (f" ({code})" if code is not None else ""))
    return "\n".join(lines)


def create(client: TeamsClient, mris: list[str], topic: str) -> str:
    members = list(dict.fromkeys([client.identity.mri, *mris]))
    if len(members) > 200:
        raise Mcp365Error("Không thể tạo nhóm trên 200 người một lần.", "Chia danh sách rồi thêm người sau.")
    props = {"threadType": "chat", "chatFilesIndexId": "2"}
    if topic:
        props["topic"] = topic
    status, raw, headers = client._chat("POST", "/threads", extra_headers={"Prefer": "return=representation"},
                                        **_wire({"properties": props, "members": [{"id": m, "role": "Admin"} for m in members]}),
                                        context="tạo nhóm Teams")
    client.invalidate_conversations()
    data = decode_json(raw, "/threads")
    location = urlsplit(headers.get("Location") or "").path
    chat = data.get("id") or (unquote(location.rsplit("/threads/", 1)[1]) if "/threads/" in location else "")
    if status == 207 or "multipleStatus" in data:
        raise Mcp365Error(f"Tạo nhóm có kết quả từng người ({chat or 'chưa có ID'}):\n" +
                          _member_status(data.get("multipleStatus", data), members), "Kiểm tra nhóm trước khi thử lại.")
    if not chat:
        raise Mcp365Error(f"Teams trả HTTP {status} nhưng không có ID nhóm.", "Kiểm tra danh sách chat, không tự tạo lại.")
    return chat


#: ``shareHistoryTime`` for "Don't include chat history": Teams Web's default choice sends -1 explicitly.
NO_HISTORY = -1


def share_history_time(days: int | None) -> int:
    """Teams Web's ``shareHistoryTime`` for an add (``ComponentsChatAddMembers`` → ``addMembersToConversation``).

    None = "Don't include chat history" → -1, sent explicitly rather than left out: Teams' readers treat
    a missing or 0 value as full history. 0 = "Include all chat history" → 0. N = "Include history from
    the past N days" → now − N days in epoch ms, as the resolver computes it (``Date.now() - N*864e5``).
    """
    if days is None:
        return NO_HISTORY
    if days < 0:
        raise Mcp365Error("share_history_days không được âm.", "Chọn None, 0 hoặc số ngày dương.")
    return 0 if days == 0 else int((time.time() - days * 86400) * 1000)


def add(client: TeamsClient, chat: str, mris: list[str], days: int | None) -> str:
    if not mris:
        raise Mcp365Error("Thiếu thành viên.", "Chọn thành viên cần thêm.")
    history = {"shareHistoryTime": share_history_time(days)}
    path = group_path(chat) + "/members"
    try:
        status, raw, _ = client._chat("PUT", path, **_wire({"members": [{"id": m, "role": "Admin", **history} for m in mris]}),
                                      context="thêm thành viên nhóm Teams")
    except Mcp365Error as exc:
        if exc.multiple_status is None:
            raise
        client.invalidate_conversations()
        raise Mcp365Error(f"HTTP {exc.http_status}: kết quả từng người:\n" + _member_status(exc.multiple_status, mris),
                          "Đọc lại danh sách thành viên trước khi thử lại.") from exc
    client.invalidate_conversations()
    data = decode_json(raw, path)
    if status == 207 or "multipleStatus" in data:
        detail = data.get("multipleStatus", data) if isinstance(data, dict) else data
        return f"HTTP {status}: kết quả từng người:\n" + _member_status(detail, mris)
    return "Đã thêm: " + ", ".join(mris)


def remove(client: TeamsClient, chat: str, mris: list[str], *, leave: bool = False) -> str:
    root = group_path(chat, meeting=leave) + "/members/"
    own = client.identity.mri.lower()
    if (leave and [m.lower() for m in mris] != [own]) or (not leave and any(m.lower() == own for m in mris)):
        raise Mcp365Error("Không thể tự xoá qua thao tác quản lý người khác.", "Dùng leave_teams_chat để rời nhóm.")
    results = run_ordered(lambda m: client._chat("DELETE", root + quote(m, safe=""), context="rời nhóm Teams" if leave else "xoá thành viên"), mris)
    client.invalidate_conversations()
    done = "đã rời nhóm" if leave else "đã xoá khỏi nhóm"
    return "\n".join(f"{m}: " + (done if err is None else f"lỗi: {err}")
                     for m, (_, err) in zip(mris, results, strict=True))


def leave(client: TeamsClient, chat: str) -> str:
    """Leave a group or meeting chat (Teams Web's Leave: delete my own roster entry). Raises when refused."""
    me = client.identity.mri
    client._chat("DELETE", group_path(chat, meeting=True) + "/members/" + quote(me, safe=""), context="rời chat Teams")
    client.invalidate_conversations()
    return "Đã rời chat"


def rename(client: TeamsClient, chat: str, topic: str) -> str:
    client._chat("PUT", group_path(chat) + "/properties?name=topic", **_wire({"topic": topic}), context="đổi tên nhóm Teams")
    client.invalidate_conversations()
    return f"Đã đổi tên `{chat}`: {topic or '(không tên)'}"


def preferences(client: TeamsClient, chat: str, mute: bool | None, hide: bool | None) -> str:
    changes = {}
    if mute is not None:
        changes["alerts"] = "false" if mute else "true"
    if hide is not None:
        stamp = str(int(time.time() * 1000)) if hide else None
        changes.update(unpinnedTime=stamp, historyHiddenTime=stamp)
    if not changes:
        raise Mcp365Error("Không có thay đổi.", "Truyền mute hoặc hide.")
    applied = []
    for prop, value in changes.items():
        try:
            client._chat("PUT", f"/users/ME/conversations/{quote(chat, safe='')}/properties?name={prop}",
                         **_wire({prop: value}), context=f"đổi {prop} của chat")
        except Mcp365Error as exc:
            if not applied:
                raise
            raise Mcp365Error(f"Đã ghi {', '.join(applied)}; lỗi ở {prop}: {exc.message}",
                              "Đọc lại trạng thái trước khi thử lại.") from exc
        applied.append(prop)
        client.invalidate_conversations()
    return "Đã cập nhật: " + ", ".join(applied)


def is_notes(chat: str) -> bool:
    return chat.strip().lower().startswith("48:")


def clear_history(client: TeamsClient, chat: str, stamp: int) -> str:
    """Teams' "Delete chat" for the signed-in account: ``clearHistoryTime`` = ``stamp`` (epoch ms, an int).

    History up to ``stamp`` disappears for you only; nobody else's view changes. Checked 19/11: the
    server refuses to set it back (``null`` / ``0`` → 400 "The clear history time is invalid").
    """
    if is_notes(chat):
        raise Mcp365Error("Không xoá chat ghi chú của bạn (48:notes) hay luồng hệ thống 48:*.", "Chọn chat 1:1 khác.")
    client._chat("PUT", f"/users/ME/conversations/{quote(chat, safe='')}/properties?name=clearHistoryTime",
                 **_wire({"clearHistoryTime": int(stamp)}), context="xoá chat phía bạn")
    client.invalidate_conversations()
    return "Đã đặt clearHistoryTime"


def delete_stamp(conv: dict, now_ms: int) -> int:
    """Teams Web ``deleteConversation``: last message time + 1, or now for a chat without messages.

    Never lower than a ``clearHistoryTime`` already set, so deleting again changes nothing old.
    """
    last = _ms((conv.get("lastMessage") or {}).get("originalarrivaltime"))
    stamp = last + 1 if last else now_ms
    return max(stamp, _ms((conv.get("properties") or {}).get("clearHistoryTime")))


def chat_party(client: TeamsClient, chat: str, conv: dict) -> str:
    """Who is on the other side of a 1:1 chat, for a delete draft: a person (name, department), a bot, or unknown.

    The thread id names a bot as ``8:orgid:<app id>`` like a person, and neither ``fetchShortProfile`` nor
    the people picker returns bots, so "not in the directory" proves nothing. Evidence, in order: the
    roster (a bot member is ``28:<app id>``), then senders of the recent messages.
    """
    peer = direct_chat_peer(chat, client.identity.mri)
    if not peer:
        return "không xác định được là người hay bot"
    bot = bot_mri(peer)
    try:
        roster = client._chat_json("GET", f"/threads/{quote(chat, safe='')}?view=msnp24Equivalent",
                                   context="xem thành viên chat 1:1")
        ids = {str(m.get("id") or "").lower() for m in roster.get("members") or [] if isinstance(m, dict)}
    except Mcp365Error:
        ids = set()
    names: dict[str, str] = {}
    messages = [conv.get("lastMessage") or {}]
    if not ({bot, peer.lower()} & ids) or not any(_sender_mri(m).lower() in (bot, peer.lower()) for m in messages):
        try:
            page = client._chat_json("GET", f"/users/ME/conversations/{quote(chat, safe='')}/messages?pageSize=50",
                                     context="xem người gửi trong chat 1:1")
            messages += page.get("messages") or []
        except Mcp365Error:
            pass
    for m in messages:
        sender = _sender_mri(m).lower()
        if sender in (bot, peer.lower()) and m.get("imdisplayname"):
            names.setdefault(sender, m["imdisplayname"])
    if bot in ids or bot in names:
        return f"bot / ứng dụng: {names.get(bot) or 'chưa rõ tên'} (`{bot}`)"
    if peer.lower() not in ids and peer.lower() not in names:
        return f"không xác định được là người hay bot (`{peer}` không còn trong danh sách thành viên, chưa thấy tin của họ)"
    try:
        rows = client._short_profiles(client._auth(), [peer])
        rows = rows.get("value", []) if isinstance(rows, dict) else rows
    except Mcp365Error:
        rows = []
    row = next((r for r in rows or [] if str(r.get("mri") or "").lower() == peer.lower()), None)
    if row:
        dept = row.get("department") or "chưa rõ đơn vị"
        return f"người: {row.get('displayName') or names.get(peer.lower()) or 'chưa rõ tên'}, {dept} (`{peer}`)"
    return (f"người: {names.get(peer.lower()) or 'chưa rõ tên'}, không tra được danh bạ "
            f"(tài khoản có thể đã khoá) (`{peer}`)")

