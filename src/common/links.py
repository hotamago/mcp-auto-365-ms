"""Biến URL trần trong văn bản thành ``<a href>`` bấm được — dùng chung cho tin Teams và mail.

Teams không tự linkify tin ``RichText/Html`` gửi qua API (người nhận phàn nàn 29/09: thấy
chữ ``https://…`` mà không bấm được). Mail HTML cũng vậy tuỳ trình đọc. Quy tắc một chỗ:

- Chỉ ``http://`` / ``https://``; trước ``http`` không được là chữ/số (``xhttps://`` bỏ qua).
- Dấu câu cuối không dính vào link: ``. , ; : ! ? ' *`` và dấu đóng/ngoặc kép Unicode;
  ``) ] }`` cuối chỉ bị bỏ khi không cân với ``( [ {`` trong URL (link Wikipedia ``Foo_(bar)``
  giữ nguyên).
- Làm trên văn bản *chưa* escape: ``href`` escape một lần (``&`` → ``&amp;``, ``"`` →
  ``&quot;``), không bao giờ escape hai lần.

:class:`Slots` giữ các mảnh HTML đã dựng (code, link) dưới dạng ký hiệu tạm ``\\x00N\\x00``
để bước escape/định dạng sau không đụng vào chúng.
"""

from __future__ import annotations

import html
import re
from collections.abc import Callable

#: URL trần, bắt rộng rồi :func:`trim_url` cắt đuôi. Dừng ở khoảng trắng, ``< > "``, dấu `
#: và ký hiệu tạm ``\x00`` của :class:`Slots`.
URL_RE = re.compile(r"(?<!\w)https?://[^\s<>\"`\x00]+", re.IGNORECASE)
#: Link markdown ``[chữ](url)``; url cho phép một lớp ngoặc cân (``…/Foo_(bar)``).
HREF_PATTERN = r"https?://(?:[^()\s\x00]|\([^()\s\x00]*\))+"

_TRAILING = set(".,;:!?'*…“”‘’«»。，、；：！？")
_CLOSERS = {")": "(", "]": "[", "}": "{", "）": "（", "】": "【", "」": "「"}
_SLOT_RE = re.compile("\x00(\\d+)\x00")


def trim_url(candidate: str) -> str:
    """Bỏ dấu câu/dấu đóng ở cuối ``candidate`` (xem đầu module)."""
    url = candidate
    while url:
        last = url[-1]
        if last in _CLOSERS:
            if url.count(_CLOSERS[last]) >= url.count(last):
                break
        elif last not in _TRAILING:
            break
        url = url[:-1]
    return url


def anchor(href: str, label_html: str | None = None) -> str:
    """``<a href="…">…</a>``; ``href`` là URL thô, ``label_html`` đã là HTML (mặc định là chính URL)."""
    label = label_html if label_html is not None else html.escape(href, quote=False)
    return f'<a href="{html.escape(href, quote=True)}">{label}</a>'


class Slots:
    """Chỗ giữ tạm các mảnh HTML đã dựng xong trong văn bản thô."""

    def __init__(self) -> None:
        self._items: list[str] = []

    def put(self, fragment: str) -> str:
        self._items.append(fragment)
        return f"\x00{len(self._items) - 1}\x00"

    def restore(self, text: str) -> str:
        # Mảnh có thể chứa mảnh khác (nhãn link có `code`): thay tới khi hết.
        for _ in range(len(self._items) + 1):
            if "\x00" not in text:
                break
            text = _SLOT_RE.sub(lambda m: self._items[int(m.group(1))], text)
        return text


def stash_urls(text: str, slots: Slots) -> str:
    """Thay mỗi URL trần trong văn bản thô bằng ký hiệu tạm của ``<a>``; đuôi dấu câu giữ lại là chữ thô."""

    def repl(m: re.Match[str]) -> str:
        raw = m.group(0)
        url = trim_url(raw)
        if not re.match(r"https?://[^/?#\s]", url, re.IGNORECASE):
            return raw
        return slots.put(anchor(url)) + raw[len(url):]

    return URL_RE.sub(repl, text)


def linkify_text(text: str, escape: Callable[[str], str] = html.escape) -> str:
    """Văn bản thường → HTML: escape mọi thứ, URL trần thành ``<a>`` (chữ hiển thị giữ nguyên)."""
    slots = Slots()
    return slots.restore(escape(stash_urls(text.replace("\x00", ""), slots)))
