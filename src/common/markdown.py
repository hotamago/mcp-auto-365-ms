"""Markdown → HTML dùng chung cho tin Teams và mail (thư viện ``markdown-it-py``).

Cấu hình (giữ hành vi cũ của hai bộ chuyển tự viết trước đây):

- Không nhận HTML thô từ input: ``a < b`` hiện đúng chữ, ``<script>`` thành ``&lt;script&gt;``.
- Xuống dòng đơn trong đoạn ⇒ ``<br>``; dòng trống tách đoạn; có danh sách ``-``/``1.``, bảng, gạch ngang.
- Link ``[chữ](url)`` chỉ nhận ``http(s)://`` và ``mailto:``; href giữ nguyên chữ (không mã hoá %), escape một lần.
- URL ``http(s)://`` trần thành link theo quy tắc chung :mod:`common.links` (Teams không tự linkify);
  trong `code` và trong nhãn link thì không.
- Tắt autolink ``<url>`` (giữ ``&lt;…&gt;`` như cũ), tắt khối code thụt 4 dấu cách và tiêu đề gạch dưới
  (``chữ\\n---``): tin chat hay thụt dòng/gạch ngang, không nên biến thành code/tiêu đề.
"""

from __future__ import annotations

import html
import re
from functools import cache

from markdown_it import MarkdownIt
from markdown_it.rules_inline import StateInline

from common.links import URL_RE, bare_url

_LINK_OK = re.compile(r"(?:https?://|mailto:)", re.IGNORECASE)
_PROTO_TAIL = re.compile(r"https?$", re.IGNORECASE)
_BLOCK_GAP = re.compile(r">\n+<")
_LIST_ITEM = re.compile(r"^\s*(?:[-*+]|\d{1,9}[.)])\s+\S")
_FENCE = re.compile(r"^\s{0,3}(```|~~~)")


def _end_lists(text: str) -> str:
    """Dòng chữ sát lề ngay sau mục danh sách là đoạn mới, không nối vào mục cuối.

    CommonMark coi đó là dòng nối lười (``- a\nKết luận`` ⇒ "Kết luận" nằm trong mục ``a``); tin chat
    và bộ chuyển cũ thì đóng danh sách. Dòng thụt lề vẫn nối vào mục; trong khối ``````` không đụng.
    """
    out: list[str] = []
    fence = ""
    in_list = False
    for line in text.replace("\r\n", "\n").split("\n"):
        m = _FENCE.match(line)
        if fence:
            fence = "" if m and m.group(1) == fence else fence
        elif m:
            fence, in_list = m.group(1), False
        elif _LIST_ITEM.match(line):
            in_list = True
        elif not line.strip():
            in_list = False
        elif in_list and not line[0].isspace():
            out.append("")
            in_list = False
        out.append(line)
    return "\n".join(out)


def _bare_url(state: StateInline, silent: bool) -> bool:
    """Luật inline: URL ``http(s)://`` trần thành link, nuốt trọn URL trước khi luật đậm/nghiêng thấy ``*``/``_``.

    Chạy ở dấu ``:`` khi chữ chờ (``pending``) kết thúc bằng ``http``/``https`` — như luật linkify của
    markdown-it nhưng theo quy tắc :mod:`common.links`. Trong `code` (luật backticks nuốt trước) và
    trong nhãn link (``linkLevel``) thì không chạy.
    """
    pos = state.pos
    if state.linkLevel > 0 or state.src[pos] != ":" or not state.src.startswith("//", pos + 1):
        return False
    proto = _PROTO_TAIL.search(state.pending)
    if not proto:
        return False
    start = pos - len(proto.group(0))
    m = URL_RE.match(state.src, start, state.posMax)
    url = bare_url(m.group(0)) if m else None
    if not url:
        return False
    if not silent:
        state.pending = state.pending[: -len(proto.group(0))]
        state.push("link_open", "a", 1).attrs = {"href": url}
        state.push("text", "", 0).content = url
        state.push("link_close", "a", -1)
    state.pos = start + len(url)
    return True


@cache
def _parser(br: str, strong: str, em: str) -> MarkdownIt:
    md = MarkdownIt("js-default", {"html": False, "breaks": True, "linkify": False, "typographer": False})
    md.disable(["autolink", "code", "lheading", "entity"])
    md.validateLink = lambda url: bool(_LINK_OK.match(url))
    md.normalizeLink = lambda url: url
    md.normalizeLinkText = lambda url: url
    md.inline.ruler.after("linkify", "bare_url", _bare_url)

    rules = md.renderer.rules
    rules["text"] = lambda tokens, idx, *_: html.escape(tokens[idx].content, quote=False)
    rules["code_inline"] = lambda tokens, idx, *_: f"<code>{html.escape(tokens[idx].content, quote=False)}</code>"
    rules["softbreak"] = rules["hardbreak"] = lambda *_: br
    rules["link_open"] = lambda tokens, idx, *_: f'<a href="{html.escape(tokens[idx].attrGet("href") or "", quote=True)}">'
    for name, tag in (("strong", strong), ("em", em)):
        rules[f"{name}_open"] = lambda *_, t=tag: f"<{t}>"
        rules[f"{name}_close"] = lambda *_, t=tag: f"</{t}>"
    return md


def markdown_to_html(text: str, *, br: str = "<br>", strong: str = "strong", em: str = "em") -> str:
    """Markdown → HTML liền (không ``\\n`` giữa các thẻ khối; trong ``<pre>`` giữ nguyên).

    ``br``/``strong``/``em`` cho từng nơi: mail ``<br>``/``<strong>``/``<em>``, Teams ``<br/>``/``<b>``/``<i>``.
    """
    rendered = _parser(br, strong, em).render(_end_lists((text or "").replace("\x00", "")))
    # Nội dung code đã escape ``<``/``>`` nên ``>\n<`` chỉ có giữa các thẻ.
    return _BLOCK_GAP.sub("><", rendered).strip()
