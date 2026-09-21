"""Microsoft Teams Chat Service client."""

from __future__ import annotations

import html as html_lib
import json
import re
import threading
import time
import unicodedata
import urllib.parse
from concurrent.futures import ThreadPoolExecutor
from datetime import UTC, datetime, timedelta, timezone
from pathlib import Path
from typing import Any

from common.config import get_config
from common.errors import ConfigError, ConversationNotFoundError, Mcp365Error, UnsupportedOperationError
from common.http import request, request_json
from common.identity import Identity, normalize_mri

from .auth import TeamsAuthManager

#: Identifier shapes that are already conversation IDs. Recognising these lets
#: get_messages() skip the conversation listing entirely - previously every
#: message fetch triggered a full list_conversations(page_size=100) call, so a
#: daily briefing issued ~25 redundant listings.
_ID_PREFIXES = ("19:", "48:", "8:orgid:", "8:live:")

_SELF_ALIASES = {"48:notes", "notes", "self", "me", "myself", "ban than", "bản thân"}

_SHAREPOINT_LINK_RE = re.compile(r'https://[a-zA-Z0-9_-]*sharepoint\.com[^\s"\'<>]+')

#: ``đ``/``Đ`` carry no combining mark, so NFD alone leaves them intact.
_EXTRA_FOLD = str.maketrans({"đ": "d", "Đ": "d", "ð": "d"})

REACTION_EMOJI = {
    "like": "👍",
    "heart": "❤️",
    "laugh": "😂",
    "surprised": "😮",
    "sad": "😢",
    "angry": "😡",
}
_REACTION_ALIASES = {
    **{name: name for name in REACTION_EMOJI},
    **{emoji: name for name, emoji in REACTION_EMOJI.items()},
    "surprise": "surprised",
}


def normalize_reaction(reaction: str) -> str:
    """Return the Chat Service reaction key accepted by Teams."""
    normalized = _REACTION_ALIASES.get((reaction or "").strip().casefold())
    if normalized:
        return normalized
    allowed = ", ".join(REACTION_EMOJI)
    raise ConfigError(
        f"Reaction Teams không hợp lệ: `{reaction}`.",
        f"Dùng một trong: {allowed}; hoặc emoji tương ứng {' '.join(REACTION_EMOJI.values())}.",
    )


def fold(text: str) -> str:
    """Casefold and strip Vietnamese diacritics for forgiving name matching.

    Chat names arrive with full diacritics (``1:1 Chat (Nguyễn Phan Nam Sơn)``)
    while people type ``nam son``. Matching the raw strings made every such
    lookup miss, so both sides go through here first.
    """
    stripped = "".join(c for c in unicodedata.normalize("NFD", text or "") if not unicodedata.combining(c))
    # Casefold first: "Ð" (U+00D0, often typed for Vietnamese "Đ") casefolds to
    # "ð", which the table maps - translating first left it as "ð" after casefold.
    return stripped.casefold().translate(_EXTRA_FOLD).strip()


def parse_attachments(raw: dict[str, Any]) -> list[dict[str, str]]:
    """Files attached to a message via the paperclip.

    They are **not** in the HTML body - a file-only message has empty content -
    but in ``properties.files``, a JSON-encoded list. Each entry points at the
    sender's OneDrive (``tenant-my.sharepoint.com/personal/...``).
    """
    files = (raw.get("properties") or {}).get("files")
    if not files:
        return []
    try:
        entries = json.loads(files) if isinstance(files, str) else files
    except (TypeError, ValueError):
        return []
    out = []
    for f in entries if isinstance(entries, list) else []:
        info = f.get("fileInfo") or {}
        url = f.get("objectUrl") or info.get("fileUrl") or ""
        if not url:
            continue
        out.append(
            {
                "name": f.get("fileName") or f.get("title") or url.rsplit("/", 1)[-1],
                "type": f.get("fileType", ""),
                "url": url,
                "share_url": info.get("shareUrl", ""),
            }
        )
    return out


def clean_teams_html(html_content: str) -> str:
    """Convert a Teams HTML message into readable text."""
    if not html_content:
        return ""

    text = html_content
    text = re.sub(r"<br\s*/?>", "\n", text)
    text = re.sub(r"</(p|div)>", "\n", text)
    text = re.sub(r"<(p|div)[^>]*>", "", text)
    text = re.sub(r'<span[^>]*itemtype="[^"]*Mention"[^>]*>([^<]*)</span>', r"@\1", text)
    text = re.sub(r'<a\s+[^>]*href="([^"]+)"[^>]*>([^<]*)</a>', r"[\2](\1)", text)
    text = re.sub(r"<[^>]+>", "", text)
    # Decode the full HTML entity set, not a hand-rolled table of five.
    text = html_lib.unescape(text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def text_to_teams_html(text: str) -> str:
    """Convert markdown-ish text to the RichText/Html Teams expects.

    Content is HTML-escaped first: an unescaped ``a < b`` used to be swallowed
    by Teams as a bogus tag.
    """
    paragraphs = (text or "").strip().split("\n\n")
    parts = []
    for para in paragraphs:
        escaped = html_lib.escape(para, quote=False)
        escaped = escaped.replace("\n", "<br/>")
        escaped = re.sub(r"\*\*([^*]+)\*\*", r"<b>\1</b>", escaped)
        escaped = re.sub(r"(?<!\*)\*([^*\n]+)\*(?!\*)", r"<i>\1</i>", escaped)
        escaped = re.sub(r"`([^`]+)`", r"<code>\1</code>", escaped)
        escaped = re.sub(r"\[([^\]]+)\]\((https?://[^)\s]+)\)", r'<a href="\2">\1</a>', escaped)
        parts.append(f"<p>{escaped}</p>")
    return "".join(parts)


_MENTION_TYPE = "http://schema.skype.com/Mention"


def apply_mentions(html_content: str, people: list[dict[str, str]]) -> tuple[str, list[dict[str, str]]]:
    """Turn ``@Name`` in the message into real Teams mentions.

    Each person is ``{"name": as written after @, "display_name", "mri"}``. The
    first ``@name`` in the text becomes a mention span; a person whose ``@name``
    is not in the text is tagged at the start of the message instead, so asking
    to tag someone never silently does nothing.

    Returns the HTML plus the ``properties.mentions`` list. ``itemid`` is the
    position in that list - Teams links the span to the entry by it.
    """
    props: list[dict[str, str]] = []
    leading: list[str] = []
    for i, person in enumerate(people):
        span = (
            f'<span itemtype="{_MENTION_TYPE}" itemscope="" itemid="{i}">'
            f"{html_lib.escape(person['display_name'], quote=False)}</span>"
        )
        token = html_lib.escape("@" + person["name"], quote=False)
        if token in html_content:
            html_content = html_content.replace(token, span, 1)
        else:
            leading.append(span)
        props.append(
            {
                "@type": _MENTION_TYPE,
                "itemid": str(i),
                "mri": person["mri"],
                "mentionType": "person",
                "displayName": person["display_name"],
            }
        )
    if leading:
        tags = " ".join(leading) + " "
        html_content = f"<p>{tags}{html_content[3:]}" if html_content.startswith("<p>") else f"<p>{tags}</p>{html_content}"
    return html_content, props


def _local_tz() -> timezone:
    return datetime.now().astimezone().tzinfo  # type: ignore[return-value]


def parse_since(since: str) -> datetime | None:
    """Interpret a ``since`` filter in the user's LOCAL timezone.

    ``"today"`` used to resolve to UTC midnight, which for a UTC+7 user silently
    dropped every message sent between 00:00 and 07:00 local time.
    """
    if not since:
        return None
    value = since.strip().lower()
    tz = _local_tz()
    now_local = datetime.now(tz)
    try:
        if value == "today":
            return now_local.replace(hour=0, minute=0, second=0, microsecond=0)
        if value == "yesterday":
            return (now_local - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
        if value.endswith("h") and value[:-1].isdigit():
            return now_local - timedelta(hours=int(value[:-1]))
        if value.endswith("d") and value[:-1].isdigit():
            return now_local - timedelta(days=int(value[:-1]))
        if len(value) == 10:
            return datetime.fromisoformat(value).replace(tzinfo=tz)
        parsed = datetime.fromisoformat(value.replace("z", "+00:00"))
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=tz)
    except ValueError:
        return None


def _parse_timestamp(value: str) -> datetime | None:
    if not value:
        return None
    try:
        text = value.replace(" ", "T") if "T" not in value else value
        text = text.replace("Z", "+00:00")
        parsed = datetime.fromisoformat(text)
        return parsed if parsed.tzinfo else parsed.replace(tzinfo=UTC)
    except ValueError:
        return None


class TeamsClient:
    """Client for the Teams Chat Service.

    Authentication is resolved lazily: constructing the client must never touch
    the keyring, otherwise an unrelated cookie problem takes down every tool in
    the server (including the SharePoint ones) at import time.
    """

    def __init__(self) -> None:
        self._conv_cache: list[dict[str, Any]] | None = None
        self._conv_cache_at: float = 0.0
        self._lock = threading.Lock()

    # ----------------------------------------------------------------- auth

    @property
    def identity(self) -> Identity:
        return TeamsAuthManager.get_identity()

    def _auth(self) -> dict[str, Any]:
        return TeamsAuthManager.get_auth()

    def _headers(self, json_body: bool = False) -> dict[str, str]:
        auth = self._auth()
        headers = {
            "Authentication": f"skypetoken={auth['token']}",
            "Accept": "application/json",
            "x-ms-client-version": "27/24081200000",
            "x-ms-client-env": "prod",
            "Origin": "https://teams.microsoft.com",
            "Referer": "https://teams.microsoft.com/",
        }
        if json_body:
            headers["Content-Type"] = "application/json"
        return headers

    # -------------------------------------------------------- conversations

    def list_conversations(
        self, page_size: int = 50, filter_keyword: str = "", use_cache: bool = True, chat_type: str = ""
    ) -> list[dict[str, Any]]:
        cfg = get_config()
        with self._lock:
            fresh = self._conv_cache is not None and (time.time() - self._conv_cache_at) < cfg.http.conversation_cache_ttl
            cached = list(self._conv_cache) if (use_cache and fresh) else None

        if cached is None:
            auth = self._auth()
            url = f"{auth['base_url']}/users/ME/conversations?view=msnp24Equivalent&pageSize={max(page_size, 50)}"
            data = request_json(url, headers=self._headers(), context="liệt kê hội thoại Teams")
            cached = [self._format_conversation(c) for c in data.get("conversations", [])]
            cached = [c for c in cached if c]
            with self._lock:
                self._conv_cache = list(cached)
                self._conv_cache_at = time.time()

        # Filter BEFORE truncating. The other order silently dropped any match
        # that sat outside the first ``page_size`` rows - a 1:1 chat 23 places
        # down the list was invisible to a keyword search with limit=15.
        results = cached
        keyword = fold(filter_keyword) if filter_keyword else ""
        if keyword:
            results = [
                c
                for c in results
                if keyword in fold(c["name"]) or keyword in fold(c.get("last_message") or "")
            ]
        if chat_type:
            wanted = fold(chat_type)
            results = [c for c in results if fold(c["type"]) == wanted]
        return results[:page_size] if page_size else results

    def _format_conversation(self, conv: dict[str, Any]) -> dict[str, Any] | None:
        conv_id = conv.get("id", "")
        if not conv_id:
            return None
        # 48:* are system notification feeds, except the personal notes chat.
        if conv_id.startswith("48:") and conv_id != "48:notes":
            return None

        props = conv.get("threadProperties", {}) or {}
        topic = props.get("topic")
        last_msg = conv.get("lastMessage", {}) or {}
        sender = last_msg.get("imdisplayname", "Unknown")

        if "@thread.tacv2" in conv_id:
            chat_type = "Channel"
        elif conv_id.startswith("19:meeting_"):
            chat_type = "MeetingChat"
        elif "@thread.v2" in conv_id:
            chat_type = "GroupChat"
        elif "@unq.gbl.spaces" in conv_id or conv_id == "48:notes":
            chat_type = "DirectChat"
        else:
            chat_type = "GroupChat"

        space_name = props.get("spaceThreadTopic")
        channel_topic = props.get("topicThreadTopic")
        if conv_id == "48:notes":
            display_name = "Chat with yourself (Notes)"
        elif space_name and channel_topic:
            display_name = f"[{space_name}] #{channel_topic}"
        elif topic:
            display_name = topic
        elif chat_type == "DirectChat":
            display_name = f"1:1 Chat ({sender})"
        else:
            display_name = conv_id

        return {
            "id": conv_id,
            "name": display_name,
            "type": chat_type,
            "last_activity": last_msg.get("composetime", ""),
            "last_sender": sender,
            "last_message": clean_teams_html(last_msg.get("content", ""))[:120].replace("\n", " "),
        }

    def find_conversation(self, identifier: str) -> dict[str, Any]:
        ident = (identifier or "").strip()
        if not ident:
            raise ConversationNotFoundError("Chưa cung cấp tên hoặc ID của cuộc trò chuyện.", "Truyền tên chat hoặc thread ID.")

        lowered = ident.lower()
        folded = fold(ident)
        if lowered in _SELF_ALIASES or lowered == self.identity.upn.lower():
            return {"id": "48:notes", "name": "Chat with yourself (Notes)", "type": "DirectChat"}

        # Already an ID: skip the listing round-trip entirely.
        if ident.startswith(_ID_PREFIXES):
            known = None
            with self._lock:
                if self._conv_cache:
                    known = next((c for c in self._conv_cache if c["id"] == ident), None)
            return known or {"id": ident, "name": ident, "type": "Unknown"}

        convs = self.list_conversations(page_size=200)
        for match in (
            lambda c: c["id"] == ident,
            lambda c: c["name"].lower() == lowered,
            lambda c: fold(c["name"]) == folded,
            lambda c: lowered in c["name"].lower(),
            lambda c: folded in fold(c["name"]),
        ):
            found = next((c for c in convs if match(c)), None)
            if found:
                return found

        available = ", ".join(c["name"] for c in convs[:8])
        raise ConversationNotFoundError(
            f"Không tìm thấy cuộc trò chuyện '{identifier}'.",
            f"Dùng `list_teams_chats` để xem danh sách. Một vài chat gần đây: {available}",
        )

    # ------------------------------------------------------------- messages

    def get_messages(
        self,
        conversation_id_or_name: str,
        limit: int = 30,
        since: str | None = None,
        only_mentions: bool = False,
        include_raw: bool = False,
    ) -> dict[str, Any]:
        auth = self._auth()
        conv = self.find_conversation(conversation_id_or_name)
        conv_id = conv["id"]

        encoded = urllib.parse.quote(conv_id)
        url = f"{auth['base_url']}/users/ME/conversations/{encoded}/messages?pageSize={max(1, min(limit, 200))}"
        data = request_json(url, headers=self._headers(), context=f"đọc tin nhắn của '{conv['name']}'")

        since_dt = parse_since(since or "")
        identity = self.identity
        formatted: list[dict[str, Any]] = []

        for raw in reversed(data.get("messages", [])):
            if raw.get("messagetype") not in ("Text", "RichText/Html"):
                continue
            if (raw.get("properties") or {}).get("deletetime"):
                continue

            compose_time = raw.get("composetime", "")
            msg_dt = _parse_timestamp(compose_time)
            if since_dt and msg_dt and msg_dt < since_dt:
                continue

            content = raw.get("content", "")
            cleaned = clean_teams_html(content)
            mentioned, reason = identity.is_mentioned(raw, cleaned)
            if only_mentions and not mentioned:
                continue

            entry = {
                "id": raw.get("id"),
                "sender": raw.get("imdisplayname", "Unknown"),
                "sender_mri": raw.get("from", "").split("/contacts/")[-1] if raw.get("from") else "",
                "timestamp": compose_time[:19].replace("T", " "),
                "timestamp_dt": msg_dt,
                "content": cleaned,
                "sharepoint_links": _SHAREPOINT_LINK_RE.findall(content),
                "attachments": parse_attachments(raw),
                "mentions_me": mentioned,
                "mention_reason": reason,
                "mentions": identity.parse_mentions(raw),
            }
            if include_raw:
                entry["raw"] = raw
            formatted.append(entry)

        return {
            "conversation_id": conv_id,
            "conversation_name": conv["name"],
            "conversation_type": conv.get("type", "Unknown"),
            "total_messages": len(formatted),
            "messages": formatted,
        }

    # ------------------------------------------------------- parallel scans

    def _scan(self, conversations: list[dict[str, Any]], worker) -> tuple[list[Any], list[str]]:
        """Run ``worker`` over conversations in parallel, collecting failures.

        Failures used to be swallowed by a bare ``except: pass``, so a throttled
        or expired session produced a *partial* feed that looked complete.
        """
        cfg = get_config().http
        errors: list[str] = []

        def guarded(conv: dict[str, Any]):
            try:
                return worker(conv)
            except Mcp365Error as exc:
                errors.append(f"{conv['name']}: {exc.message}")
            except Exception as exc:  # noqa: BLE001 - surfaced to the caller
                errors.append(f"{conv['name']}: {type(exc).__name__}: {exc}")
            return None

        workers = max(1, min(cfg.max_workers, len(conversations) or 1))
        with ThreadPoolExecutor(max_workers=workers) as pool:
            results = list(pool.map(guarded, conversations))
        return [r for r in results if r is not None], errors

    def _active_conversations(self, types: tuple[str, ...], limit: int, keyword: str = "") -> list[dict[str, Any]]:
        convs = self.list_conversations(page_size=200, filter_keyword=keyword)
        return [c for c in convs if c["type"] in types][:limit]

    def get_recent_feed(
        self, hours: int = 48, max_chats: int = 8, limit_per_chat: int = 8, filter_keyword: str = ""
    ) -> dict[str, Any]:
        chats = self._active_conversations(("GroupChat", "MeetingChat", "Channel"), max_chats, filter_keyword)
        cutoff = datetime.now(UTC) - timedelta(hours=hours)

        def worker(conv):
            res = self.get_messages(conv["id"], limit=limit_per_chat)
            recent = [m for m in res["messages"] if not m["timestamp_dt"] or m["timestamp_dt"] >= cutoff]
            if not recent:
                return None
            return {
                "chat_name": conv["name"],
                "chat_id": conv["id"],
                "chat_type": conv["type"],
                "last_activity": conv["last_activity"],
                "messages": recent,
            }

        feed, errors = self._scan(chats, worker)
        feed.sort(key=lambda f: f.get("last_activity") or "", reverse=True)
        return {"feed": feed, "errors": errors, "scanned": len(chats)}

    def get_user_mentions(
        self,
        hours: int = 72,
        limit: int = 20,
        context_before: int = 2,
        context_after: int = 2,
        max_chats: int = 25,
    ) -> dict[str, Any]:
        """Find mentions of the signed-in user across chats, channels and 1:1s."""
        chats = self._active_conversations(("GroupChat", "MeetingChat", "Channel", "DirectChat"), max_chats)
        cutoff = datetime.now(UTC) - timedelta(hours=hours)

        def worker(conv):
            res = self.get_messages(conv["id"], limit=max(30, limit * 2))
            msgs = res["messages"]
            hits = []
            for idx, msg in enumerate(msgs):
                if msg["timestamp_dt"] and msg["timestamp_dt"] < cutoff:
                    continue
                if not msg["mentions_me"]:
                    continue
                start, end = max(0, idx - context_before), min(len(msgs), idx + context_after + 1)
                hits.append(
                    {
                        "chat_name": conv["name"],
                        "chat_id": conv["id"],
                        "chat_type": conv["type"],
                        "message_id": msg["id"],
                        "sender": msg["sender"],
                        "timestamp": msg["timestamp"],
                        "content": msg["content"],
                        "mention_reason": msg["mention_reason"],
                        "sharepoint_links": msg["sharepoint_links"],
                        "context": [
                            {
                                "sender": msgs[j]["sender"],
                                "timestamp": msgs[j]["timestamp"],
                                "content": msgs[j]["content"],
                                "sharepoint_links": msgs[j]["sharepoint_links"],
                                "is_mention": j == idx,
                                "offset": j - idx,
                            }
                            for j in range(start, end)
                        ],
                    }
                )
            return hits or None

        groups, errors = self._scan(chats, worker)
        mentions = [m for group in groups for m in group]
        mentions.sort(key=lambda m: m["timestamp"], reverse=True)
        return {"mentions": mentions[:limit], "errors": errors, "scanned": len(chats)}

    def get_new_mentions_since(self, cursor: str = "", limit: int = 20) -> dict[str, Any]:
        """Cursor-based mention polling, for use from a scheduled loop.

        A stdio MCP server cannot hold a long-running watch loop, so the caller
        keeps the cursor and polls.
        """
        since_dt = parse_since(cursor) if cursor else None
        hours = 24
        if since_dt:
            delta = datetime.now(since_dt.tzinfo) - since_dt
            hours = max(1, min(int(delta.total_seconds() // 3600) + 1, 24 * 14))

        result = self.get_user_mentions(hours=hours, limit=limit)
        if since_dt:
            filtered = []
            for mention in result["mentions"]:
                ts = _parse_timestamp(mention["timestamp"])
                if ts is None or ts > since_dt:
                    filtered.append(mention)
            result["mentions"] = filtered

        newest = max((m["timestamp"] for m in result["mentions"]), default="")
        result["cursor"] = newest or (cursor or datetime.now(UTC).strftime("%Y-%m-%d %H:%M:%S"))
        return result

    def search_messages(self, keywords: list[str] | str, max_results: int = 20, max_chats: int = 20) -> dict[str, Any]:
        terms = [keywords] if isinstance(keywords, str) else list(keywords or [])
        terms = [str(t).lower().strip() for t in terms if str(t).strip()]
        if not terms:
            raise Mcp365Error("Chưa cung cấp từ khoá tìm kiếm.", "Truyền ít nhất một từ khoá, ví dụ ['DTC', 'CAN'].")

        chats = self._active_conversations(("GroupChat", "MeetingChat", "Channel", "DirectChat"), max_chats)

        def worker(conv):
            res = self.get_messages(conv["id"], limit=30)
            hits = []
            for msg in res["messages"]:
                low = msg["content"].lower()
                matched = next((t for t in terms if t in low), None)
                if matched:
                    hits.append({**msg, "chat_name": conv["name"], "chat_id": conv["id"], "matched_keyword": matched})
            return hits or None

        groups, errors = self._scan(chats, worker)
        hits = [h for group in groups for h in group]
        hits.sort(key=lambda h: h["timestamp"], reverse=True)
        return {"hits": hits[:max_results], "errors": errors, "scanned": len(chats)}

    # -------------------------------------------------------------- sending

    def resolve_mentions(self, conversation_id_or_name: str, names: list[str], scan: int = 200) -> list[dict[str, str]]:
        """Find who to tag, by name, among people seen in this conversation.

        The Chat Service has no people search, but every message carries its
        sender's MRI and every mention carries the mentioned person's MRI. So a
        name resolves if that person has written or been tagged in the recent
        history - which covers anyone the user would realistically tag there.
        Matching ignores diacritics and case, and must be unambiguous.
        """
        conv = self.find_conversation(conversation_id_or_name)
        # Every name a person has appeared under. Whoever types a tag can shorten
        # it ("Hoàng" instead of "Đỗ Văn Hoàng (…)"), so keeping only the first
        # name seen made full-name lookups miss people who had plainly written in
        # the chat. The sender name is the canonical display.
        aliases: dict[str, set[str]] = {}
        full_name: dict[str, str] = {}
        for msg in self.get_messages(conv["id"], limit=scan)["messages"]:
            if msg.get("sender") and msg.get("sender_mri"):
                mri = normalize_mri(msg["sender_mri"])
                aliases.setdefault(mri, set()).add(msg["sender"])
                full_name.setdefault(mri, msg["sender"])
            for tagged in msg.get("mentions") or []:
                if tagged.get("mri") and tagged.get("displayName"):
                    aliases.setdefault(normalize_mri(tagged["mri"]), set()).add(tagged["displayName"])
        seen = {mri: full_name.get(mri) or max(known, key=len) for mri, known in aliases.items()}

        people = []
        for name in names:
            wanted = fold(name.lstrip("@"))
            hits = {
                mri: seen[mri] for mri, known in aliases.items() if wanted and any(wanted in fold(n) for n in known)
            }
            if len(hits) == 1:
                mri, display = next(iter(hits.items()))
                people.append({"name": name.lstrip("@"), "display_name": display, "mri": mri})
            elif not hits:
                raise ConversationNotFoundError(
                    f"Không tìm thấy '{name}' trong {scan} tin gần nhất của '{conv['name']}'.",
                    "Chỉ tag được người đã nhắn hoặc đã được tag trong chat này. Kiểm tra lại tên "
                    "(có thể viết không dấu), hoặc để người đó nhắn một lần trước.",
                )
            else:
                raise Mcp365Error(
                    f"'{name}' khớp nhiều người: " + "; ".join(sorted(hits.values())),
                    "Ghi đầy đủ họ tên hơn để chỉ còn đúng một người.",
                )
        return people

    def _build_quote(self, conv_id: str, reply_to_id: str) -> str:
        sender, preview = "Member", ""
        try:
            for msg in self.get_messages(conv_id, limit=50)["messages"]:
                if str(msg.get("id")) == str(reply_to_id):
                    sender = msg.get("sender", "Member")
                    preview = msg.get("content", "")[:150]
                    break
        except Mcp365Error:
            pass
        return (
            f'<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="{html_lib.escape(str(reply_to_id))}">'
            f'<strong itemprop="mri">{html_lib.escape(sender)}</strong>'
            f'<span itemprop="time" itemid="{html_lib.escape(str(reply_to_id))}"></span>'
            f'<p itemprop="preview">{html_lib.escape(preview)}</p>'
            f"</blockquote>"
        )

    def send_message(
        self,
        conversation_id_or_name: str,
        message: str,
        reply_to_id: str | None = None,
        file_path: str | None = None,
        mentions: list[dict[str, str]] | None = None,
    ) -> dict[str, Any]:
        """Send a message. ``mentions`` are people from :meth:`resolve_mentions`."""
        auth = self._auth()
        conv = self.find_conversation(conversation_id_or_name)
        conv_id, conv_name = conv["id"], conv["name"]

        file_info = None
        if file_path:
            # NOTE: pathlib was previously not imported in this module, so every
            # attachment upload raised NameError before reaching SharePoint.
            local = Path(file_path).expanduser().resolve()
            if not local.is_file():
                raise Mcp365Error(
                    f"Không tìm thấy file đính kèm: {file_path}",
                    "Kiểm tra lại đường dẫn (dùng đường dẫn tuyệt đối nếu cần).",
                )
            from sharepoint.client import SharePointClient

            try:
                file_info = SharePointClient().upload_file(
                    str(local), target_folder_url_or_path=get_config().sharepoint.attachment_folder
                )
            except Mcp365Error as exc:
                # Say plainly that nothing went out, or the agent cannot tell
                # whether a retry would post the message twice.
                raise Mcp365Error(
                    f"Tin nhắn CHƯA được gửi: upload file đính kèm '{local.name}' lên SharePoint thất bại.\n"
                    f"{exc.message}",
                    exc.remediation,
                ) from exc
            size = file_info["size"]
            size_str = f"{size / (1024 * 1024):.2f} MB" if size > 1024 * 1024 else f"{size / 1024:.1f} KB"
            message = f"{message}\n\n📎 **Tệp đính kèm:** [{file_info['name']}]({file_info['webUrl']}) *({size_str})*"

        html_content = text_to_teams_html(message)
        mention_props: list[dict[str, str]] = []
        if mentions:
            html_content, mention_props = apply_mentions(html_content, mentions)
        if reply_to_id:
            html_content = self._build_quote(conv_id, reply_to_id) + html_content

        client_message_id = str(int(time.time() * 1000))
        payload: dict[str, Any] = {
            "content": html_content,
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "clientmessageid": client_message_id,
            "imdisplayname": self.identity.display_name or "Unknown",
        }
        if mention_props:
            # Teams expects the list JSON-encoded inside properties, not nested.
            payload["properties"] = {"mentions": json.dumps(mention_props, ensure_ascii=False)}
        url = f"{auth['base_url']}/users/ME/conversations/{urllib.parse.quote(conv_id)}/messages"
        data = request_json(
            url,
            headers=self._headers(json_body=True),
            method="POST",
            data=json.dumps(payload).encode("utf-8"),
            context=f"gửi tin nhắn tới '{conv_name}'",
        )
        return {
            "status": "SENT",
            "conversation_id": conv_id,
            "conversation_name": conv_name,
            "message_id": self._resolve_sent_id(conv_id, client_message_id, data),
            "message_sent": message,
            "reply_to_id": reply_to_id,
            "attached_file": file_info,
            "mentioned": [p["display_name"] for p in mentions or []],
        }

    def _resolve_sent_id(self, conv_id: str, client_message_id: str, response: dict[str, Any]) -> str:
        """Find the server-side id of a message we just sent.

        The send endpoint returns only ``OriginalArrivalTime``, so the id is
        recovered by matching our ``clientmessageid`` in the recent history;
        the arrival time is used as a fallback because in practice it equals
        the message id.
        """
        try:
            recent = self.get_messages(conv_id, limit=15, include_raw=True)
            for msg in reversed(recent["messages"]):
                if str((msg.get("raw") or {}).get("clientmessageid", "")) == client_message_id:
                    return str(msg["id"])
        except Mcp365Error:
            pass
        arrival = response.get("OriginalArrivalTime")
        return str(arrival) if arrival else ""

    def reply_to_channel_thread(self, channel_name_or_id: str, parent_message_id: str, message: str) -> dict[str, Any]:
        """Post inside an existing channel thread rather than starting a new one.

        Teams channels address a thread with a ``;messageid=<root>`` suffix on
        the conversation id; ``send_message`` alone always starts a new thread.
        """
        auth = self._auth()
        conv = self.find_conversation(channel_name_or_id)
        if "@thread.tacv2" not in conv["id"]:
            raise UnsupportedOperationError(
                f"'{conv['name']}' không phải là Teams Channel.",
                "Với group chat thường, dùng `send_teams_message` kèm `reply_to_id` để trích dẫn.",
            )

        thread_id = f"{conv['id']};messageid={parent_message_id}"
        payload = {
            "content": text_to_teams_html(message),
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "clientmessageid": str(int(time.time() * 1000)),
            "imdisplayname": self.identity.display_name or "Unknown",
        }
        url = f"{auth['base_url']}/users/ME/conversations/{urllib.parse.quote(thread_id)}/messages"
        data = request_json(
            url,
            headers=self._headers(json_body=True),
            method="POST",
            data=json.dumps(payload).encode("utf-8"),
            context=f"trả lời thread trong '{conv['name']}'",
        )
        return {
            "status": "REPLIED",
            "conversation_id": conv["id"],
            "conversation_name": conv["name"],
            "thread_id": thread_id,
            "parent_message_id": parent_message_id,
            "message_id": data.get("id"),
            "message_sent": message,
        }

    def edit_message(self, conversation_id_or_name: str, message_id: str, new_message: str) -> dict[str, Any]:
        auth = self._auth()
        conv = self.find_conversation(conversation_id_or_name)
        url = (
            f"{auth['base_url']}/users/ME/conversations/{urllib.parse.quote(conv['id'])}"
            f"/messages/{urllib.parse.quote(str(message_id))}"
        )
        request(
            url,
            headers=self._headers(json_body=True),
            method="PUT",
            data=json.dumps(
                {"content": text_to_teams_html(new_message), "messagetype": "RichText/Html", "contenttype": "text"}
            ).encode("utf-8"),
            context=f"sửa tin nhắn trong '{conv['name']}'",
        )
        return {
            "status": "EDITED",
            "conversation_id": conv["id"],
            "conversation_name": conv["name"],
            "message_id": message_id,
            "new_message": new_message,
        }

    def delete_message(self, conversation_id_or_name: str, message_id: str) -> dict[str, Any]:
        auth = self._auth()
        conv = self.find_conversation(conversation_id_or_name)
        url = (
            f"{auth['base_url']}/users/ME/conversations/{urllib.parse.quote(conv['id'])}"
            f"/messages/{urllib.parse.quote(str(message_id))}"
        )
        request(url, headers=self._headers(), method="DELETE", context=f"xoá tin nhắn trong '{conv['name']}'")
        return {
            "status": "DELETED",
            "conversation_id": conv["id"],
            "conversation_name": conv["name"],
            "message_id": message_id,
        }

    def react_to_message(
        self,
        conversation_id_or_name: str,
        message_id: str,
        reaction: str,
        *,
        remove: bool = False,
    ) -> dict[str, Any]:
        """Add or remove the signed-in user's reaction on one message."""
        reaction_key = normalize_reaction(reaction)
        auth = self._auth()
        conv = self.find_conversation(conversation_id_or_name)
        url = (
            f"{auth['base_url']}/users/ME/conversations/{urllib.parse.quote(conv['id'], safe='')}"
            f"/messages/{urllib.parse.quote(str(message_id), safe='')}/properties?name=emotions"
        )
        emotion: dict[str, Any] = {"key": reaction_key}
        if not remove:
            emotion["value"] = int(time.time() * 1000)
        payload = {"emotions": json.dumps(emotion, separators=(",", ":"))}
        headers = self._headers(json_body=True)
        headers["x-ms-client-caller"] = (
            "updateMessageReactionRemove" if remove else "updateMessageReactionAdd"
        )
        request(
            url,
            headers=headers,
            method="DELETE" if remove else "PUT",
            data=json.dumps(payload).encode("utf-8"),
            context=f"{'gỡ' if remove else 'thả'} reaction trong '{conv['name']}'",
        )
        return {
            "status": "REMOVED" if remove else "REACTED",
            "conversation_id": conv["id"],
            "conversation_name": conv["name"],
            "message_id": str(message_id),
            "reaction": reaction_key,
            "emoji": REACTION_EMOJI[reaction_key],
        }

    # ------------------------------------------------------------ calendar

    def get_calendar_events(self, start: datetime, end: datetime) -> list[dict[str, Any]]:
        """Calendar via the Teams middle tier.

        The Graph token from the Azure CLI carries no ``Calendars.*`` scope, so
        Graph is not an option here. The middle-tier bearer token from the
        browser session is accepted, but the events path is undocumented and has
        changed between Teams releases - hence ``teams.calendar_endpoint`` in
        the config file.
        """
        auth = self._auth()
        token = auth.get("middle_tier_token")
        if not token:
            raise UnsupportedOperationError(
                "Không tìm thấy 'authtoken' của Teams trong Chrome (cần cho Lịch).",
                "Mở https://teams.microsoft.com trong Chrome và đăng nhập, rồi thử lại.",
            )
        if auth.get("middle_tier_exp", 0) and auth["middle_tier_exp"] <= time.time():
            raise UnsupportedOperationError(
                "Token middle-tier của Teams đã hết hạn.",
                "Tải lại tab https://teams.microsoft.com trong Chrome.",
            )

        cfg = get_config().teams
        region = cfg.middle_tier_region or auth["region"]
        path = cfg.calendar_endpoint.format(
            region=region,
            start=start.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.000Z"),
            end=end.astimezone(UTC).strftime("%Y-%m-%dT%H:%M:%S.000Z"),
        )
        data = request_json(
            f"https://teams.microsoft.com{path}",
            headers={"Authorization": f"Bearer {token}", "Accept": "application/json"},
            context="đọc lịch từ Teams middle-tier",
        )
        events = data.get("value") or data.get("events") or (data if isinstance(data, list) else [])
        out = []
        for ev in events if isinstance(events, list) else []:
            out.append(
                {
                    "subject": ev.get("subject") or ev.get("title") or "(không tiêu đề)",
                    "start": (ev.get("startTime") or ev.get("start") or ""),
                    "end": (ev.get("endTime") or ev.get("end") or ""),
                    "organizer": (ev.get("organizer") or {}).get("displayName")
                    if isinstance(ev.get("organizer"), dict)
                    else ev.get("organizer", ""),
                    "is_online": bool(ev.get("isOnlineMeeting") or ev.get("skypeTeamsMeetingUrl")),
                    "join_url": ev.get("skypeTeamsMeetingUrl") or ev.get("onlineMeetingJoinUrl") or "",
                    "location": ev.get("location") or "",
                }
            )
        return out
