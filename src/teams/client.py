"""Microsoft Teams Chat Service client."""

from __future__ import annotations

import contextvars
import html as html_lib
import json
import logging
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
from common.http import decode_json, request, request_json, request_to_file
from common.identity import Identity, normalize_mri

from . import endpoints
from .auth import TeamsAuthManager
from .endpoints import is_teams_media_url, split_chat_url

logger = logging.getLogger(__name__)

#: Identifier shapes that are already conversation IDs. Recognising these lets
#: get_messages() skip the conversation listing entirely - previously every
#: message fetch triggered a full list_conversations(page_size=100) call, so a
#: daily briefing issued ~25 redundant listings.
_ID_PREFIXES = ("19:", "48:", "8:orgid:", "8:live:")

_SELF_ALIASES = {"48:notes", "notes", "self", "me", "myself", "ban than", "bản thân"}

_SHAREPOINT_LINK_RE = re.compile(r'https://[a-zA-Z0-9_-]*sharepoint\.com[^\s"\'<>]+')

#: ``đ``/``Đ`` carry no combining mark, so NFD alone leaves them intact.
_EXTRA_FOLD = str.maketrans({"đ": "d", "Đ": "d", "ð": "d"})

_IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".gif", ".webp", ".bmp", ".svg")
_IMG_TAG_RE = re.compile(r"<img\s+([^>]+)>", re.IGNORECASE)

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


def normalize_direct_chat_id(conv_id: str) -> str:
    """Normalize a 1:1 direct chat space ID by lexicographically sorting the GUIDs.

    Teams Chat Service requires the two GUIDs in a 1:1 conversation thread
    (19:{guid1}_{guid2}@unq.gbl.spaces) to be lexicographically sorted.
    If passed reversed, the service responds with:
    404 LocationLookupFailed for thread ...
    """
    val = (conv_id or "").strip()
    if val.startswith("19:") and val.endswith("@unq.gbl.spaces"):
        inner = val[3:-15]
        if "_" in inner:
            parts = inner.split("_", 1)
            if len(parts[0]) > 10 and len(parts[1]) > 10:
                g1, g2 = sorted([parts[0].lower(), parts[1].lower()])
                return f"19:{g1}_{g2}@unq.gbl.spaces"
    return val


def direct_chat_peer(conv_id: str, my_mri: str) -> str:
    """The MRI of the *other* member of a 1:1 chat, read off the thread id.

    ``19:{guid1}_{guid2}@unq.gbl.spaces`` names both members, so who the chat is
    with never has to be guessed from whoever happened to send the last message.
    Guessing was the bug: a 1:1 chat where the signed-in user spoke last was
    labelled with the signed-in user's own name.

    Returns ``""`` when the id is not a 1:1 thread, when both GUIDs are the same
    (a chat with oneself) or when neither GUID is the signed-in user - the
    caller then keeps the old last-sender label.
    """
    val = (conv_id or "").strip()
    if not (val.startswith("19:") and val.endswith("@unq.gbl.spaces")):
        return ""
    parts = [p for p in val[3:-15].split("_") if p]
    if len(parts) != 2:
        return ""
    first, second = parts[0].lower(), parts[1].lower()
    if first == second:
        return ""
    me = (my_mri or "").rsplit(":", 1)[-1].lower()
    if not me:
        return ""
    if me == first:
        return normalize_mri(second)
    if me == second:
        return normalize_mri(first)
    return ""


def _sender_mri(message: dict[str, Any]) -> str:
    """The sender's MRI from a message's ``from`` link, normalised."""
    link = message.get("from") or ""
    return normalize_mri(link.split("/contacts/")[-1]) if link else ""


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


def parse_inline_images(html_content: str) -> list[dict[str, str]]:
    """Extract inline images/screenshots embedded in Teams HTML message content.

    Ignores emojis/emoticons (itemtype=".../Emoji" or ".../Emoticon").
    """
    if not html_content or "<img" not in html_content:
        return []
    out = []
    idx = 1
    for match in _IMG_TAG_RE.finditer(html_content):
        attrs = match.group(1)
        if 'itemtype="http://schema.skype.com/Emoji"' in attrs or 'itemtype="http://schema.skype.com/Emoticon"' in attrs:
            continue
        src_m = re.search(r'src=["\']([^"\']+)["\']', attrs)
        if not src_m:
            continue
        url = src_m.group(1)
        obj_m = re.search(r"/objects/([^/]+)/", url)
        img_id = obj_m.group(1) if obj_m else ""
        if not img_id:
            id_m = re.search(r'id=["\']([^"\']+)["\']', attrs)
            img_id = id_m.group(1) if id_m else f"img_{idx}"

        ext = ".png" if "png" in attrs.lower() or "png" in url.lower() else ".jpg"
        name = f"image_{img_id[:12]}{ext}"
        out.append({"id": img_id, "url": url, "name": name, "type": "inline"})
        idx += 1
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
    def _replace_img(m: re.Match) -> str:
        attrs = m.group(1)
        if 'itemtype="http://schema.skype.com/Emoji"' in attrs or 'itemtype="http://schema.skype.com/Emoticon"' in attrs:
            return ""
        return " 🖼️ [image] "

    text = _IMG_TAG_RE.sub(_replace_img, text)
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
_REPLY_TYPE = "http://schema.skype.com/Reply"
#: A reply quote nested in the quoted message (quoting a reply). Teams leaves it
#: out of the preview; it used to leak in as "NameText..." run together.
_NESTED_QUOTE_RE = re.compile(
    r"<blockquote\b[^>]*schema\.skype\.com/Reply[^>]*>.*?</blockquote>", re.IGNORECASE | re.DOTALL
)
#: Teams cuts the preview at 199 characters and appends an ellipsis.
_PREVIEW_CHARS = 199


def quote_preview(html_content: str) -> str:
    """Plain-text preview of a message, as the Teams client puts in a reply quote."""
    text = _NESTED_QUOTE_RE.sub(" ", html_content or "")
    text = re.sub(r"<br\s*/?>|</(p|div|li|h\d)>", " ", text, flags=re.IGNORECASE)
    text = _IMG_TAG_RE.sub(" ", text)
    text = html_lib.unescape(re.sub(r"<[^>]+>", "", text))
    text = re.sub(r"\s+", " ", text).strip()
    return text if len(text) <= _PREVIEW_CHARS else text[:_PREVIEW_CHARS] + "…"


def _epoch_ms(value: str) -> int | None:
    parsed = _parse_timestamp(value)
    return int(parsed.timestamp()) * 1000 + parsed.microsecond // 1000 if parsed else None


def build_reply_quote(original: dict[str, Any]) -> tuple[str, dict[str, Any]]:
    """The two halves of a real Teams quote reply to ``original`` (a raw message).

    Taken from replies the Teams client itself sent (self chat and a 1:1, 22/09):

    * the HTML ``<blockquote itemtype=".../Reply" itemid="<msg id>">`` whose
      ``<strong itemprop="mri">`` carries the author's MRI in ``itemid``;
    * ``properties.qtdMsgs``, a JSON *string* ``[{"messageId", "sender",
      "time"}]``. The server validates it (it comes back with
      ``validationResult: "Valid"`` and ``hasValidMsgReferences``); it is what
      ties the quote to the quoted message.

    The old payload had the blockquote only, without the author's MRI and
    without ``qtdMsgs``: Teams showed a styled block, not a reply.
    """
    msg_id = str(original.get("id") or "")
    sender = (original.get("from") or "").split("/contacts/")[-1]
    name = original.get("imdisplayname") or sender
    sent_at = _epoch_ms(original.get("originalarrivaltime") or original.get("composetime") or "")
    if sent_at is None and msg_id.isdigit():
        sent_at = int(msg_id)  # server message ids are the arrival time in ms
    esc_id = html_lib.escape(msg_id)
    html_quote = (
        f'<blockquote itemscope="" itemtype="{_REPLY_TYPE}" itemid="{esc_id}">\r\n'
        f'<strong itemprop="mri" itemid="{html_lib.escape(sender)}">{html_lib.escape(name, quote=False)}</strong>'
        f'<span itemprop="time" itemid="{esc_id}"></span>\r\n'
        f'<p itemprop="preview">{html_lib.escape(quote_preview(original.get("content", "")), quote=False)}</p>\r\n'
        f"</blockquote>\r\n"
    )
    return html_quote, {"messageId": msg_id, "sender": sender, "time": sent_at}


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


def render_message(message: str, people: list[dict[str, str]] | None) -> tuple[str, dict[str, str]]:
    """Message text -> ``(html, properties)`` as send and edit both post it.

    The one place mentions are built: ``properties.mentions`` goes out
    JSON-encoded inside ``properties``, not nested, as the Teams client does.
    """
    html_content = text_to_teams_html(message)
    properties: dict[str, str] = {}
    if people:
        html_content, mention_props = apply_mentions(html_content, people)
        properties["mentions"] = json.dumps(mention_props, ensure_ascii=False)
    return html_content, properties


_MENTION_SPAN_RE = re.compile(
    r'<span[^>]*itemtype="http://schema\.skype\.com/Mention"[^>]*itemid="(\d+)"[^>]*>(.*?)</span>', re.S
)
_ORG_SUFFIX_RE = re.compile(r"\s*\([^()]*\)\s*$")


def kept_mentions(original: dict[str, Any], new_message: str) -> tuple[list[dict[str, str]], list[str]]:
    """People tagged in ``original`` whose ``@Name`` is still written in ``new_message``.

    A tag is kept when the new text contains ``@`` + its display name, the text
    shown in its span, or either without the trailing ``(Org unit)``; the
    longest form written wins, so the whole ``@Name (Org)`` becomes the tag.
    Returns ``(people, dropped display names)``; ``people`` is in the shape
    :func:`apply_mentions` takes.
    """
    props = original.get("properties") or {}
    raw = props.get("mentions") or []
    if isinstance(raw, str):
        try:
            raw = json.loads(raw)
        except json.JSONDecodeError:
            raw = []
    shown = {
        idx: html_lib.unescape(re.sub(r"<[^>]+>", "", text)).strip()
        for idx, text in _MENTION_SPAN_RE.findall(original.get("content") or "")
    }
    people: list[dict[str, str]] = []
    dropped: list[str] = []
    seen: set[str] = set()
    for m in raw if isinstance(raw, list) else []:
        mri = normalize_mri(str(m.get("mri") or ""))
        display = str(m.get("displayName") or "").strip()
        if not mri or not display or mri in seen:
            continue
        if str(m.get("mentionType") or "person").lower() != "person":
            dropped.append(display)  # team/channel tags: not rebuilt here
            continue
        seen.add(mri)
        forms = {display, shown.get(str(m.get("itemid")), "")}
        forms |= {_ORG_SUFFIX_RE.sub("", f) for f in forms}
        written = [f for f in sorted(forms, key=len, reverse=True) if f and f"@{f}" in new_message]
        if written:
            people.append({"name": written[0], "display_name": display, "mri": mri})
        else:
            dropped.append(display)
    return people, dropped


def _transport(url: str, **kwargs: Any) -> tuple[int, bytes, Any]:
    """The HTTP call the endpoint router makes.

    ``request`` is looked up in this module at call time, so tests that patch
    ``teams.client.request`` intercept Chat Service traffic, probes included.
    """
    return request(url, **kwargs)


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



#: ``messagetype`` values that are messages a person typed (as opposed to system events).
_CHAT_MESSAGE_TYPES = ("Text", "RichText", "RichText/Html")

class TeamsClient:
    """Client for the Teams Chat Service.

    Authentication is resolved lazily: constructing the client must never touch
    the keyring, otherwise an unrelated cookie problem takes down every tool in
    the server (including the SharePoint ones) at import time.
    """

    def __init__(self) -> None:
        self._conv_cache: list[dict[str, Any]] | None = None
        self._conv_cache_at: float = 0.0
        # MRI -> display name, learned from every conversation listing. A 1:1
        # chat payload carries no roster names, so the only free source of a
        # colleague's name is a message they sent - in *any* chat on the page.
        self._peer_names: dict[str, str] = {}
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

    def _chat(
        self,
        method: str,
        path_or_url: str,
        *,
        data: bytes | None = None,
        json_body: bool = False,
        extra_headers: dict[str, str] | None = None,
        context: str = "",
    ) -> tuple[int, bytes, Any]:
        """One Chat Service request via the endpoint router (fastest host, failover).

        ``path_or_url`` is relative to the ``/v1`` base (``/users/ME/...``) or an
        absolute URL the service handed back, which is re-routed as well.
        """
        auth = self._auth()
        headers = self._headers(json_body=json_body)
        if extra_headers:
            headers.update(extra_headers)
        return endpoints.ROUTER.call(
            auth.get("region") or "apac",
            path_or_url,
            method=method,
            headers=headers,
            data=data,
            context=context,
            transport=_transport,
        )

    def _chat_json(self, method: str, path_or_url: str, **kwargs: Any) -> dict:
        _status, body, _headers = self._chat(method, path_or_url, **kwargs)
        return decode_json(body, path_or_url)

    # -------------------------------------------------------- conversations
    def create_or_get_direct_chat(self, target: str) -> str:
        """Ensure a 1:1 conversation thread exists with the target user.

        ``target`` can be:
        - A direct chat space ID (e.g. ``19:{guid1}_{guid2}@unq.gbl.spaces``)
        - An MRI (e.g. ``8:orgid:<guid>``)
        - A user GUID

        Calls ``POST /threads`` on the Teams Chat Service to provision the thread
        if it does not exist yet, and returns the canonical conversation ID.
        """
        target = (target or "").strip()
        auth = self._auth()
        my_mri = auth["identity"].mri
        my_guid = my_mri.removeprefix("8:orgid:").lower()

        target_mri = ""
        if target.startswith("8:orgid:"):
            target_mri = target
        elif target.startswith("19:") and "@unq.gbl.spaces" in target:
            inner = target.removeprefix("19:").removesuffix("@unq.gbl.spaces")
            if "_" in inner:
                parts = inner.split("_", 1)
                p1, p2 = parts[0].lower(), parts[1].lower()
                other_guid = p2 if p1 == my_guid else p1
                target_mri = f"8:orgid:{other_guid}"
        elif len(target) == 36 and target.count("-") == 4:
            target_mri = f"8:orgid:{target.lower()}"

        if not target_mri:
            return normalize_direct_chat_id(target)

        target_guid = target_mri.removeprefix("8:orgid:").lower()
        if target_guid == my_guid:
            return "48:notes"

        body = {
            "members": [
                {"id": my_mri, "role": "Admin"},
                {"id": target_mri, "role": "Admin"},
            ],
            "properties": {
                "threadType": "chat",
                "chatFilesIndexId": "2",
                "uniquerosterthread": "true",
                "fixedRoster": "true",
            },
        }
        try:
            _status, _resp_body, resp_headers = self._chat(
                "POST",
                "/threads",
                data=json.dumps(body).encode("utf-8"),
                json_body=True,
                context=f"khởi tạo cuộc trò chuyện 1:1 với {target_mri}",
            )
            location = resp_headers.get("Location") or ""
            if "/threads/" in location:
                return location.rsplit("/threads/", 1)[-1]
        except Exception as exc:
            logger.warning("Không thể khởi tạo thread 1:1 qua POST /threads (%s); dùng ID chuẩn hoá", exc)

        g1, g2 = sorted([my_guid, target_guid])
        return f"19:{g1}_{g2}@unq.gbl.spaces"


    def list_conversations(
        self, page_size: int = 50, filter_keyword: str = "", use_cache: bool = True, chat_type: str = ""
    ) -> list[dict[str, Any]]:
        cfg = get_config()
        with self._lock:
            fresh = self._conv_cache is not None and (time.time() - self._conv_cache_at) < cfg.http.conversation_cache_ttl
            cached = list(self._conv_cache) if (use_cache and fresh) else None

        if cached is None:
            data = self._chat_json(
                "GET",
                f"/users/ME/conversations?view=msnp24Equivalent&pageSize={max(page_size, 50)}",
                context="liệt kê hội thoại Teams",
            )
            raw_convs = data.get("conversations", [])
            my_mri = self.identity.mri
            # Pass 1 learns names, pass 2 formats: a 1:1 chat where I spoke last
            # borrows the peer's name from wherever they did speak last.
            names = self._learn_peer_names(raw_convs, my_mri)
            cached = [self._format_conversation(c, my_mri=my_mri, peer_names=names) for c in raw_convs]
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

    def _learn_peer_names(self, raw_convs: list[dict[str, Any]], my_mri: str) -> dict[str, str]:
        """Remember ``MRI -> display name`` for everyone but me, and return the map.

        The cache lives on the client (``tools.teams()`` is a singleton), so a
        name seen on an earlier page still labels a chat whose peer is silent on
        this one.

        Only real chat messages teach a name. A system event as ``lastMessage``
        (``ThreadActivity/AddMember``, call logs, topic updates) can carry an
        ``imdisplayname`` that is not the sender's own display name.
        """
        learned: dict[str, str] = {}
        for conv in raw_convs:
            last_msg = conv.get("lastMessage") or {}
            if last_msg.get("messagetype") not in _CHAT_MESSAGE_TYPES:
                continue
            mri = _sender_mri(last_msg)
            name = last_msg.get("imdisplayname") or ""
            if mri and name and mri != my_mri:
                learned[mri] = name
        with self._lock:
            self._peer_names.update(learned)
            return dict(self._peer_names)

    def _format_conversation(
        self,
        conv: dict[str, Any],
        my_mri: str | None = None,
        peer_names: dict[str, str] | None = None,
    ) -> dict[str, Any] | None:
        conv_id = conv.get("id", "")
        if not conv_id:
            return None
        # 48:* are system notification feeds, except the personal notes chat.
        if conv_id.startswith("48:") and conv_id != "48:notes":
            return None

        if my_mri is None:
            my_mri = self.identity.mri
        if peer_names is None:
            with self._lock:
                peer_names = dict(self._peer_names)

        props = conv.get("threadProperties", {}) or {}
        topic = props.get("topic")
        last_msg = conv.get("lastMessage", {}) or {}
        sender = last_msg.get("imdisplayname", "Unknown")
        sender_mri = _sender_mri(last_msg)

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
            display_name = f"1:1 Chat ({self._peer_label(conv_id, my_mri, sender, sender_mri, peer_names)})"
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

    @staticmethod
    def _peer_label(
        conv_id: str, my_mri: str, sender: str, sender_mri: str, peer_names: dict[str, str]
    ) -> str:
        """Name the other person in a 1:1 chat - never the signed-in user.

        The last sender is only a valid label when the last sender *is* the peer.
        Otherwise the name comes from the learned map, and failing that from the
        peer's own MRI: an unhelpful label beats a wrong one.
        """
        peer = direct_chat_peer(conv_id, my_mri)
        if not peer:
            # Not a two-party thread id (a bot chat, a self chat, an id shape we
            # do not know): the last sender is the best available label.
            return sender
        if sender_mri and sender_mri == peer and sender:
            return sender
        return peer_names.get(peer) or peer

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
            if "@unq.gbl.spaces" in ident:
                ident = normalize_direct_chat_id(ident)
            elif ident.startswith(("8:orgid:", "8:live:")):
                direct_id = self.create_or_get_direct_chat(ident)
                return {"id": direct_id, "name": f"1:1 Chat ({ident})", "type": "DirectChat"}

            known = None
            with self._lock:
                if self._conv_cache:
                    known = next((c for c in self._conv_cache if c["id"] == ident), None)
            return known or {"id": ident, "name": ident, "type": "DirectChat" if "@unq.gbl.spaces" in ident else "Unknown"}

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

        # Not found in existing conversations: search company directory
        try:
            users = self.search_users(identifier, max_results=3)
            if users:
                top = users[0]
                top_name = fold(top.get("name") or "")
                top_email = (top.get("email") or "").lower()
                top_upn = (top.get("upn") or "").lower()
                if (
                    len(users) == 1
                    or folded == top_name
                    or lowered == top_email
                    or lowered == top_upn
                    or folded in top_name
                ):
                    target_mri = top.get("teams_mri") or ""
                    if target_mri:
                        direct_id = self.create_or_get_direct_chat(target_mri)
                        return {
                            "id": direct_id,
                            "name": f"1:1 Chat ({top['name']})",
                            "type": "DirectChat",
                        }
        except Exception as exc:
            logger.debug("Không thể tìm người trong danh bạ cho '%s': %s", identifier, exc)

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
        conv = self.find_conversation(conversation_id_or_name)
        conv_id = conv["id"]

        encoded = urllib.parse.quote(conv_id)
        path = f"/users/ME/conversations/{encoded}/messages?pageSize={max(1, min(limit, 200))}"
        try:
            data = self._chat_json("GET", path, context=f"đọc tin nhắn của '{conv['name']}'")
        except Mcp365Error as err:
            if "@unq.gbl.spaces" in conv_id and ("404" in str(err) or "LocationLookupFailed" in str(err)):
                canonical_id = self.create_or_get_direct_chat(conv_id)
                if canonical_id != conv_id:
                    conv_id = canonical_id
                    encoded = urllib.parse.quote(conv_id)
                    path = f"/users/ME/conversations/{encoded}/messages?pageSize={max(1, min(limit, 200))}"
                    try:
                        data = self._chat_json("GET", path, context=f"đọc tin nhắn của '{conv['name']}'")
                    except Mcp365Error:
                        return {"conversation_id": conv_id, "conversation_name": conv["name"], "conversation_type": conv.get("type", "DirectChat"), "messages": [], "count": 0}
                else:
                    return {"conversation_id": conv_id, "conversation_name": conv["name"], "conversation_type": conv.get("type", "DirectChat"), "messages": [], "count": 0}
            else:
                raise

        since_dt = parse_since(since or "")
        identity = self.identity
        formatted: list[dict[str, Any]] = []

        for raw in reversed(data.get("messages", [])):
            if raw.get("messagetype") not in _CHAT_MESSAGE_TYPES:
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

            attachments = parse_attachments(raw)
            inline_imgs = parse_inline_images(content)
            attachment_imgs = [
                {
                    "id": a["url"].rsplit("/", 1)[-1],
                    "url": a["url"],
                    "name": a["name"],
                    "type": "attachment",
                }
                for a in attachments
                if any(a["name"].lower().endswith(ext) for ext in _IMAGE_EXTS)
            ]
            entry = {
                "id": raw.get("id"),
                "sender": raw.get("imdisplayname", "Unknown"),
                "sender_mri": raw.get("from", "").split("/contacts/")[-1] if raw.get("from") else "",
                "timestamp": compose_time[:19].replace("T", " "),
                "timestamp_dt": msg_dt,
                "content": cleaned,
                "sharepoint_links": _SHAREPOINT_LINK_RE.findall(content),
                "attachments": attachments,
                "images": inline_imgs + attachment_imgs,
                "mentions_me": mentioned,
                "mention_reason": reason,
                "mentions": identity.parse_mentions(raw),
            }
            if include_raw:
                entry["raw"] = raw
            formatted.append(entry)

        # Reading a chat is the other free source of ``MRI -> name``: it makes
        # the next listing label a 1:1 chat properly even when the peer has been
        # silent on the conversations page.
        my_mri = identity.mri
        with self._lock:
            for entry in formatted:
                mri = normalize_mri(entry.get("sender_mri") or "")
                name = entry.get("sender") or ""
                if mri and mri != my_mri and name and name != "Unknown":
                    self._peer_names[mri] = name

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
        # Pool threads start with an empty context; carry the tool's
        # ``timeout_seconds`` into them. One Context cannot be entered by two
        # threads at once, hence a copy per task.
        parent = contextvars.copy_context()
        with ThreadPoolExecutor(max_workers=workers) as pool:
            results = list(pool.map(lambda conv: parent.copy().run(guarded, conv), conversations))
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
    def download_image(self, url: str, target_file_path: Path | str) -> Path:
        """Download an inline Teams AMS image or attachment image to a local file.

        An image served from a Chat Service base goes through the endpoint
        router, like any other chat call, so a dead front door in the link does
        not doom the download. Other Teams media hosts (AMS, async gateway, the
        web-app proxies) are recognised by host, not by a substring anywhere in
        the URL.
        """
        target = Path(target_file_path).expanduser().resolve()
        target.parent.mkdir(parents=True, exist_ok=True)

        if split_chat_url(url):
            auth = self._auth()
            _status, data, _headers = self._chat(
                "GET",
                url,
                extra_headers={"Cookie": f"skypetoken_asm={auth['token']}", "Accept": "*/*"},
                context=f"tải ảnh Teams '{target.name}'",
            )
            target.write_bytes(data)
            return target

        if is_teams_media_url(url):
            auth = self._auth()
            headers = {
                "Cookie": f"skypetoken_asm={auth['token']}",
                "User-Agent": get_config().http.user_agent,
                "Accept": "*/*",
            }
            host = (urllib.parse.urlsplit(url).hostname or "").lower()
            if host == "teams.microsoft.com" or host.endswith("teams.cloud.microsoft"):
                # The web-app proxies authenticate like the Chat Service itself.
                headers["Authentication"] = f"skypetoken={auth['token']}"
            request_to_file(url, target, headers=headers, context=f"tải ảnh Teams '{target.name}'")
            return target

        if "sharepoint.com" in url:
            from sharepoint.client import SharePointClient

            sp = SharePointClient()
            request_to_file(url, target, headers=sp._download_headers(url), context=f"tải ảnh đính kèm '{target.name}'")
            return target

        request_to_file(url, target, context=f"tải ảnh '{target.name}'")
        return target

    def download_message_images(
        self,
        conversation_id_or_name: str,
        message_id: str = "",
        target_dir: str = "",
        limit: int = 10,
    ) -> list[dict[str, Any]]:
        """Download images from a specific message or recent messages in a chat."""
        conv = self.find_conversation(conversation_id_or_name)
        dest_dir = Path(target_dir).expanduser().resolve() if target_dir else Path("downloads/images").resolve()
        dest_dir.mkdir(parents=True, exist_ok=True)

        res = self.get_messages(conv["id"], limit=50)
        messages = res["messages"]
        if message_id:
            messages = [m for m in messages if str(m.get("id")) == str(message_id)]
            if not messages:
                raise ConversationNotFoundError(
                    f"Không tìm thấy message ID '{message_id}' trong 50 tin gần nhất của '{conv['name']}'.",
                    "Kiểm tra lại message ID hoặc tăng phạm vi quét.",
                )

        downloaded: list[dict[str, Any]] = []
        for msg in reversed(messages):
            images = msg.get("images") or []
            for img in images:
                if len(downloaded) >= limit:
                    break
                file_name = f"{msg['id']}_{img['name']}"
                target_path = dest_dir / file_name
                try:
                    self.download_image(img["url"], target_path)
                    downloaded.append({
                        "message_id": msg["id"],
                        "sender": msg["sender"],
                        "timestamp": msg["timestamp"],
                        "name": file_name,
                        "path": str(target_path),
                        "size": target_path.stat().st_size,
                        "type": img.get("type", "image"),
                    })
                except Exception as exc:
                    logger.warning("Could not download image %s: %s", img["url"], exc)

            if len(downloaded) >= limit:
                break

        return downloaded


    def search_users(self, query: str, max_results: int = 10) -> list[dict[str, Any]]:
        """Search for colleagues across the organization directory.

        Searches via Outlook Web People API using the browser session. Returns
        rich profiles with display name, email, UPN, title, department, phone,
        Teams MRI (8:orgid:<guid>) and direct 1:1 chat ID.
        """
        q = (query or "").strip()
        if not q:
            return []

        my_guid = ""
        try:
            my_mri = self.identity.mri or ""
            guid_match = re.search(r"([0-9a-fA-F-]{36})", my_mri)
            if guid_match:
                my_guid = guid_match.group(1).lower()
        except Exception:
            pass

        limit = max(1, min(max_results, 50))
        people_list: list[dict[str, Any]] = []

        # 1. Primary channel: Outlook Web People Search API
        try:
            from outlook.auth import MailAuthManager

            auth = MailAuthManager()
            token = auth.get_token()
            headers = {"Authorization": f"Bearer {token}", "Accept": "application/json"}
            encoded = urllib.parse.quote(f'"{q}"')
            url = f"https://outlook.office.com/api/v2.0/me/people?$top={limit}&$search={encoded}"
            res = request_json(url, headers=headers, context=f"tìm kiếm người '{q}'")
            for item in res.get("value", []):
                raw_id = item.get("Id", "")
                guid_match = re.search(r"([0-9a-fA-F-]{36})", raw_id)
                object_id = guid_match.group(1).lower() if guid_match else ""
                mri = f"8:orgid:{object_id}" if object_id else ""
                direct_chat_id = ""
                if my_guid and object_id and my_guid != object_id:
                    g1, g2 = sorted([my_guid.lower(), object_id.lower()])
                    direct_chat_id = f"19:{g1}_{g2}@unq.gbl.spaces"

                emails = [e.get("Address") for e in item.get("ScoredEmailAddresses", []) if e.get("Address")]
                if not emails:
                    emails = [e.get("Address") for e in item.get("EmailAddresses", []) if e.get("Address")]

                phones = [p.get("Number") for p in item.get("Phones", []) if p.get("Number")]

                people_list.append({
                    "name": item.get("DisplayName") or "",
                    "given_name": item.get("GivenName") or "",
                    "surname": item.get("Surname") or "",
                    "email": emails[0] if emails else "",
                    "all_emails": emails,
                    "upn": item.get("UserPrincipalName") or "",
                    "job_title": item.get("JobTitle") or "",
                    "department": item.get("Department") or "",
                    "office": item.get("OfficeLocation") or "",
                    "phone": phones[0] if phones else "",
                    "object_id": object_id,
                    "teams_mri": mri,
                    "direct_chat_id": direct_chat_id,
                })
        except Exception as exc:
            logger.info("Outlook People Search API unavailable (%s); checking conversation roster", exc)

        if people_list:
            return people_list

        # 2. Fallback channel: Search recent conversations roster
        wanted = fold(q.lstrip("@"))
        seen_mris: set[str] = set()
        try:
            convs = self.list_conversations(page_size=50)
            for c in convs:
                c_name = c.get("name") or ""
                if c.get("chat_type") == "DirectChat" and wanted and wanted in fold(c_name):
                    p_name = c_name
                    if "(" in c_name and c_name.endswith(")"):
                        p_name = c_name[c_name.find("(") + 1 : -1].strip()
                    c_id = c.get("id") or ""
                    other_guid = ""
                    guid_matches = re.findall(r"([0-9a-fA-F-]{36})", c_id)
                    for g in guid_matches:
                        if g.lower() != my_guid:
                            other_guid = g.lower()
                            break
                    mri = f"8:orgid:{other_guid}" if other_guid else ""
                    if mri and mri not in seen_mris:
                        seen_mris.add(mri)
                        people_list.append({
                            "name": p_name,
                            "given_name": "",
                            "surname": "",
                            "email": "",
                            "all_emails": [],
                            "upn": "",
                            "job_title": "",
                            "department": "",
                            "office": "",
                            "phone": "",
                            "object_id": other_guid,
                            "teams_mri": mri,
                            "direct_chat_id": c_id,
                        })
        except Exception as exc:
            logger.debug("Conversation roster search error: %s", exc)

        return people_list[:limit]

    def resolve_mentions(self, conversation_id_or_name: str, names: list[str], scan: int = 200) -> list[dict[str, str]]:
        """Find who to tag, by name, among people seen in this conversation or company directory."""
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
                searched = self.search_users(name.lstrip("@"), max_results=3)
                if len(searched) == 1:
                    p = searched[0]
                    people.append({"name": name.lstrip("@"), "display_name": p["name"], "mri": p["teams_mri"]})
                elif len(searched) > 1:
                    raise Mcp365Error(
                        f"'{name}' không có trong lịch sử chat và khớp nhiều người trong danh bạ: "
                        + "; ".join(p["name"] for p in searched),
                        "Ghi đầy đủ họ tên hoặc email để tag chính xác.",
                    )
                else:
                    raise ConversationNotFoundError(
                        f"Không tìm thấy '{name}' trong lịch sử chat của '{conv['name']}' hoặc danh bạ tổ chức.",
                        "Kiểm tra lại tên hoặc nhập email/alias để tìm.",
                    )
            else:
                raise Mcp365Error(
                    f"'{name}' khớp nhiều người: " + "; ".join(sorted(hits.values())),
                    "Ghi đầy đủ họ tên hơn để chỉ còn đúng một người.",
                )
        return people

    def _get_raw_message(self, conv_id: str, message_id: str) -> dict[str, Any]:
        """One message as the Chat Service stores it (author MRI, arrival time, HTML)."""
        encoded = urllib.parse.quote(conv_id)
        try:
            return self._chat_json(
                "GET", f"/users/ME/conversations/{encoded}/messages/{urllib.parse.quote(message_id)}",
                context=f"đọc tin nhắn gốc {message_id}",
            )
        except Mcp365Error as exc:
            # Fall back to the recent history, in case the single-message route is refused.
            try:
                history = self.get_messages(conv_id, limit=200, include_raw=True)["messages"]
            except Mcp365Error:
                raise exc from None
            for msg in history:
                if str(msg.get("id")) == message_id:
                    return msg["raw"]
            raise exc

    def _reply_quote(self, conv_id: str, reply_to_id: str) -> tuple[str, dict[str, Any]]:
        """Quote HTML + ``qtdMsgs`` entry for a reply (see :func:`build_reply_quote`).

        Fails instead of degrading: the old code, when it could not find the
        message, still sent a quote attributed to "Member" with no preview.
        """
        reply_to_id = str(reply_to_id).strip()
        try:
            original = self._get_raw_message(conv_id, reply_to_id)
        except Mcp365Error as exc:
            raise Mcp365Error(
                f"Tin nhắn CHƯA được gửi: không đọc được tin gốc {reply_to_id} để trích dẫn.\n{exc.message}",
                "Kiểm tra `reply_to_id` là id tin nhắn (lấy từ `read_teams_chat`) trong đúng cuộc trò chuyện này, "
                "hoặc gửi lại không kèm `reply_to_id`.",
            ) from exc
        if str(original.get("id") or "") != reply_to_id or (original.get("properties") or {}).get("deletetime"):
            raise Mcp365Error(
                f"Tin nhắn CHƯA được gửi: tin gốc {reply_to_id} không còn (đã bị xoá hoặc không thuộc chat này).",
                "Chọn tin nhắn khác để trả lời, hoặc gửi không kèm `reply_to_id`.",
            )
        return build_reply_quote(original)

    def send_message(
        self,
        conversation_id_or_name: str,
        message: str,
        reply_to_id: str | None = None,
        file_path: str | None = None,
        mentions: list[dict[str, str]] | None = None,
    ) -> dict[str, Any]:
        """Send a message. ``mentions`` are people from :meth:`resolve_mentions`."""
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

        html_content, properties = render_message(message, mentions)
        if reply_to_id:
            quote_html, quoted = self._reply_quote(conv_id, reply_to_id)
            html_content = quote_html + html_content
            # Same encoding as the Teams client: a compact JSON string.
            properties["qtdMsgs"] = json.dumps([quoted], ensure_ascii=False, separators=(",", ":"))
            properties["formatVariant"] = "TEAMS"

        client_message_id = str(int(time.time() * 1000))
        payload: dict[str, Any] = {
            "content": html_content,
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "clientmessageid": client_message_id,
            "imdisplayname": self.identity.display_name or "Unknown",
        }
        if properties:
            payload["properties"] = properties
        try:
            data = self._chat_json(
                "POST",
                f"/users/ME/conversations/{urllib.parse.quote(conv_id)}/messages",
                data=json.dumps(payload).encode("utf-8"),
                json_body=True,
                context=f"gửi tin nhắn tới '{conv_name}'",
            )
        except Mcp365Error as err:
            if "@unq.gbl.spaces" in conv_id and ("404" in str(err) or "LocationLookupFailed" in str(err)):
                logger.info("Thread 1:1 '%s' chưa khởi tạo, tự động gọi create_or_get_direct_chat...", conv_id)
                canonical_id = self.create_or_get_direct_chat(conv_id)
                conv_id = canonical_id
                data = self._chat_json(
                    "POST",
                    f"/users/ME/conversations/{urllib.parse.quote(conv_id)}/messages",
                    data=json.dumps(payload).encode("utf-8"),
                    json_body=True,
                    context=f"gửi tin nhắn tới '{conv_name}'",
                )
            else:
                raise
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
        data = self._chat_json(
            "POST",
            f"/users/ME/conversations/{urllib.parse.quote(thread_id)}/messages",
            data=json.dumps(payload).encode("utf-8"),
            json_body=True,
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

    def mentions_to_keep(
        self, conversation_id_or_name: str, message_id: str, new_message: str
    ) -> tuple[list[dict[str, str]], list[str]]:
        """Tags of the message being edited that its new text still writes as ``@Name``.

        See :func:`kept_mentions`. Fails instead of guessing when the original
        cannot be read: editing blind would silently turn tags into plain text.
        """
        conv = self.find_conversation(conversation_id_or_name)
        try:
            original = self._get_raw_message(conv["id"], str(message_id))
        except Mcp365Error as exc:
            raise Mcp365Error(
                f"Tin nhắn CHƯA được sửa: không đọc được tin gốc {message_id} để giữ các tag cũ.\n{exc.message}",
                "Kiểm tra `message_id`, hoặc truyền `mentions` (danh sách tên) / `mentions=[]` (bỏ hết tag).",
            ) from exc
        return kept_mentions(original, new_message)

    def edit_message(
        self,
        conversation_id_or_name: str,
        message_id: str,
        new_message: str,
        mentions: list[dict[str, str]] | None = None,
    ) -> dict[str, Any]:
        """Replace the text of one of our messages.

        ``mentions`` are people from :meth:`resolve_mentions`; ``None`` keeps
        the original's tags still written as ``@Name`` (:meth:`mentions_to_keep`),
        ``[]`` tags nobody. Without this the PUT carried only plain HTML and
        every ``@Name`` of the original became ordinary text.
        """
        conv = self.find_conversation(conversation_id_or_name)
        if mentions is None:
            mentions, _dropped = self.mentions_to_keep(conv["id"], message_id, new_message)
        html_content, properties = render_message(new_message, mentions)
        payload: dict[str, Any] = {"content": html_content, "messagetype": "RichText/Html", "contenttype": "text"}
        if properties:
            payload["properties"] = properties
        path = (
            f"/users/ME/conversations/{urllib.parse.quote(conv['id'])}"
            f"/messages/{urllib.parse.quote(str(message_id))}"
        )
        self._chat(
            "PUT",
            path,
            data=json.dumps(payload).encode("utf-8"),
            json_body=True,
            context=f"sửa tin nhắn trong '{conv['name']}'",
        )
        return {
            "status": "EDITED",
            "conversation_id": conv["id"],
            "conversation_name": conv["name"],
            "message_id": message_id,
            "new_message": new_message,
            "mentioned": [p["display_name"] for p in mentions],
        }

    def delete_message(self, conversation_id_or_name: str, message_id: str) -> dict[str, Any]:
        conv = self.find_conversation(conversation_id_or_name)
        path = (
            f"/users/ME/conversations/{urllib.parse.quote(conv['id'])}"
            f"/messages/{urllib.parse.quote(str(message_id))}"
        )
        self._chat("DELETE", path, context=f"xoá tin nhắn trong '{conv['name']}'")
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
        conv = self.find_conversation(conversation_id_or_name)
        path = (
            f"/users/ME/conversations/{urllib.parse.quote(conv['id'], safe='')}"
            f"/messages/{urllib.parse.quote(str(message_id), safe='')}/properties?name=emotions"
        )
        emotion: dict[str, Any] = {"key": reaction_key}
        if not remove:
            emotion["value"] = int(time.time() * 1000)
        payload = {"emotions": json.dumps(emotion, separators=(",", ":"))}
        self._chat(
            "DELETE" if remove else "PUT",
            path,
            data=json.dumps(payload).encode("utf-8"),
            json_body=True,
            extra_headers={
                "x-ms-client-caller": "updateMessageReactionRemove" if remove else "updateMessageReactionAdd"
            },
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
