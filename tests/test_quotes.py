"""Tin được trích dẫn / chuyển tiếp: hiện tin gốc (id, người gửi, giờ, chat, chữ, file) và tải file của nó.

Dữ liệu dựng lại từ tin thật đọc về ngày 28/09 (cùng cấu trúc HTML + ``properties``), đã ẩn danh:
MRI, id chat, tên người, tên file và link đều là giả.
"""

from __future__ import annotations

import json
from urllib.parse import unquote

import pytest

from common.errors import Mcp365Error
from teams import client as client_mod
from teams.client import TeamsClient, parse_quoted, quoted_line

ME = "8:orgid:00000000-0000-0000-0000-0000000000a1"
PEER = "8:orgid:00000000-0000-0000-0000-0000000000b2"
THIRD = "8:orgid:00000000-0000-0000-0000-0000000000c3"
CHAT = "19:00000000-0000-0000-0000-0000000000a1_00000000-0000-0000-0000-0000000000b2@unq.gbl.spaces"
OTHER_CHAT = "19:0000000000000000000000000000abcd@thread.v2"
LOCKED_CHAT = "19:0000000000000000000000000000dead@thread.v2"
_CONTACT = "https://apac.ng.msg.teams.microsoft.com/v1/users/ME/contacts/"


def _file(name: str) -> dict:
    """Một mục ``properties.files`` như Teams ghi khi gửi file qua kẹp giấy."""
    url = f"https://contoso-my.sharepoint.com/personal/user_contoso_com/Documents/Microsoft%20Teams%20Chat%20Files/{name}"
    return {
        "itemid": "11111111-2222-3333-4444-555555555555",
        "fileName": name,
        "fileType": name.rsplit(".", 1)[-1],
        "fileInfo": {"itemId": None, "fileUrl": url, "shareUrl": url + "?share=1"},
        "@type": "http://schema.skype.com/File",
        "objectUrl": url,
        "title": name,
        "state": "active",
    }


def _msg(msg_id: str, sender: str, name: str, content: str, **props) -> dict:
    return {
        "id": msg_id,
        "messagetype": "RichText/Html",
        "from": _CONTACT + sender,
        "imdisplayname": name,
        "composetime": "2026-09-28T07:22:55.4860000Z",
        "originalarrivaltime": "2026-09-28T07:22:55.4860000Z",
        "content": content,
        "properties": {"files": "[]", **props},
    }


def _qtd(msg_id: str, sender: str) -> list[dict]:
    """``qtdMsgs`` như đọc về: ``message``, ``sharedRefId``, ``replyChainId`` luôn null."""
    return [{"messageId": int(msg_id), "sender": sender, "time": int(msg_id), "message": None,
             "validationResult": "Valid", "sharedRefId": None, "replyChainId": None}]


def _reply_html(msg_id: str, author: str, name: str, preview: str) -> str:
    return (
        f'<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="{msg_id}">'
        f'<strong itemprop="mri" itemid="{author}">{name}</strong>'
        f'<span itemprop="time" itemid="{msg_id}"></span><p itemprop="preview">{preview}</p></blockquote>\n'
    )


FILE_NAMES = ("huong-dan.md", "schema.md", "logic.md", "main.md")
#: Tin gốc chỉ có 4 file, không chữ (22/09).
FILES_ONLY = {
    **_msg("1790062855722", ME, "Người Dùng (VF-TEST)", ""),
    "composetime": "2026-09-22T07:40:55.7220000Z",
    "originalarrivaltime": "2026-09-22T07:40:55.7220000Z",
}
FILES_ONLY["properties"]["files"] = json.dumps([_file(n) for n in FILE_NAMES])
#: Tin trích tin trên: preview chỉ còn "📄 📄 📄 📄", không id chat, không tên file.
QUOTES_FILES = _msg(
    "1790580175486", PEER, "Bạn Chat (VF-TEST)",
    _reply_html("1790062855722", ME, "Người Dùng (VF-TEST)", "📄 📄 📄 📄")
    + "<p>vẫn dùng 4 file này đúng không</p>",
    qtdMsgs=_qtd("1790062855722", ME), hasValidMsgReferences=True,
)
TEXT_ORIGINAL = _msg("1790500000001", PEER, "Bạn Chat (VF-TEST)", "<p>chiều nay họp lúc 3h nhé</p>")
QUOTES_TEXT = _msg(
    "1790500000002", ME, "Người Dùng (VF-TEST)",
    _reply_html("1790500000001", PEER, "Bạn Chat (VF-TEST)", "chiều nay họp lúc 3h nhé") + "<p>ok</p>",
    qtdMsgs=_qtd("1790500000001", PEER),
)
#: Hai tin chuyển tiếp từ nhóm khác; khối đầu có một trích dẫn lồng (tin ở nhóm gốc).
FORWARD = _msg(
    "1790418269081", ME, "Người Dùng (VF-TEST)",
    '<blockquote itemtype="http://schema.skype.com/Forward">'
    + _reply_html("1790417975375", PEER, "Bạn Chat (VF-TEST)", " ")
    + '<p>câu thứ nhất</p></blockquote><blockquote itemtype="http://schema.skype.com/Forward">'
    "<p>câu thứ hai</p></blockquote>",
    files=json.dumps([_file("plan.xlsx")]),
    forwardTemplateId="basic_forward_message_template",
    originalMessageContext={
        "originalSender": THIRD, "originalSentTime": "2026-09-26T10:22:32.847Z", "messageId": 1790418152847,
        "originalThreadId": OTHER_CHAT, "threadType": "chat", "messageType": "RichText/Html", "files": "[]",
    },
    originalMessageContext1={
        "originalSender": THIRD, "originalSentTime": "2026-09-26T10:22:38.143Z", "messageId": 1790418158143,
        "originalThreadId": OTHER_CHAT, "threadType": "chat", "messageType": "RichText/Html",
    },
    additionalMessageContext={"amsreferences": [], "cards": "[]", "links": "[]"},
)
FORWARDED_ORIGINAL = _msg("1790418152847", THIRD, "Người Thứ Ba (VF-TEST)", "<p>câu thứ nhất</p>")
NESTED_ORIGINAL = _msg("1790417975375", PEER, "Bạn Chat (VF-TEST)", "<p>tin trong nhóm gốc</p>",
                       files=json.dumps([_file("log.txt")]))
#: Khối trích dẫn không có id, không có qtdMsgs (gặp ở tin do bot/ứng dụng khác gửi).
NO_ID = _msg(
    "1790500000009", PEER, "Bạn Chat (VF-TEST)",
    '<blockquote itemscope="" itemtype="http://schema.skype.com/Reply">'
    '<strong itemprop="mri">Ai Đó</strong><p itemprop="preview">📄</p></blockquote><p>ok</p>',
)


class FakeChat:
    """Chat Service giả: trang tin theo chat, tin lẻ theo (chat, id); chat bị khoá trả 403."""

    def __init__(self, pages: dict[str, list[dict]], store: dict[tuple[str, str], dict] | None = None,
                 locked: tuple[str, ...] = ()) -> None:
        self.pages, self.store, self.locked = pages, store or {}, set(locked)
        self.single: list[tuple[str, str]] = []

    def __call__(self, method: str, path: str, **_kw) -> dict:
        path = unquote(path)
        chat = path.split("/conversations/", 1)[1].split("/messages", 1)[0]
        if "/messages?" in path:
            return {"messages": list(reversed(self.pages.get(chat, [])))}  # API: mới nhất trước
        msg_id = path.rsplit("/", 1)[-1]
        self.single.append((chat, msg_id))
        if chat in self.locked:
            raise Mcp365Error("Bị từ chối truy cập (HTTP 403) khi đọc tin được trích.", "Không có quyền.")
        for page_chat, msgs in self.pages.items():
            for raw in msgs:
                if page_chat == chat and raw["id"] == msg_id:
                    return raw
        if (chat, msg_id) in self.store:
            return self.store[(chat, msg_id)]
        raise Mcp365Error(f"HTTP 404 Not Found khi đọc tin {msg_id}.", "Kiểm tra id.")


@pytest.fixture
def chat(monkeypatch, identity):
    c = TeamsClient()
    monkeypatch.setattr(type(c), "identity", property(lambda self: identity))
    monkeypatch.setattr(c, "list_conversations", lambda *a, **k: [])

    def install(pages, store=None, locked=()):
        fake = FakeChat(pages, store, locked)
        monkeypatch.setattr(c, "_chat_json", fake)
        return fake

    return c, install


def _by_id(res: dict, msg_id: str) -> dict:
    return next(m for m in res["messages"] if str(m["id"]) == msg_id)


# ------------------------------------------------------------------ dữ liệu nhúng


def test_embedded_quote_has_id_author_time_but_no_files():
    (item,) = parse_quoted(QUOTES_FILES)
    assert item["kind"] == "reply" and item["message_id"] == "1790062855722"
    assert item["sender_mri"] == ME and item["sender_name"] == "Người Dùng (VF-TEST)"
    assert item["timestamp"] == "2026-09-22 07:40:55"
    assert item["conversation_id"] == "" and item["preview"] == "📄 📄 📄 📄"
    assert item["attachments"] == [] and item["resolved"] is False


def test_forward_blocks_map_to_original_message_context_in_order():
    forward1, nested, forward2 = parse_quoted(FORWARD)
    assert (forward1["kind"], forward1["message_id"], forward1["conversation_id"]) == (
        "forward", "1790418152847", OTHER_CHAT)
    assert forward1["sender_mri"] == THIRD and forward1["timestamp"] == "2026-09-26 10:22:32"
    assert forward2["message_id"] == "1790418158143" and forward2["preview"] == "câu thứ hai"
    # Trích dẫn lồng trong khối chuyển tiếp trỏ vào chat GỐC, không phải chat đang đọc.
    assert (nested["kind"], nested["message_id"], nested["conversation_id"]) == (
        "reply", "1790417975375", OTHER_CHAT)
    assert nested["sender_name"] == "Bạn Chat (VF-TEST)"


def test_quote_without_id_keeps_author_and_preview():
    (item,) = parse_quoted(NO_ID)
    assert item["message_id"] == "" and item["sender_name"] == "Ai Đó" and item["preview"] == "📄"
    assert quoted_line(item) == '↩️ trích tin của Ai Đó: "📄"'


def test_author_in_a_span_is_read_too():
    raw = _msg(
        "2", PEER, "B",
        '<div><blockquote itemscope="" itemtype="http://schema.skype.com/Reply" itemid="1790215601557">'
        f'<p><b><span itemprop="mri" itemid="{PEER}" style="font-size:small">Bạn Chat (VF-TEST)&nbsp;</span>'
        '<span itemprop="time" itemid="1790215601557"></span></b></p>'
        '<p itemprop="preview">&#128247; Bảng</p></blockquote>C ơi</div>',
    )
    (item,) = parse_quoted(raw)
    assert (item["sender_name"], item["sender_mri"], item["preview"]) == ("Bạn Chat (VF-TEST)", PEER, "📷 Bảng")


def test_qtd_msgs_without_html_block_still_listed():
    raw = _msg("3", PEER, "B", "<p>chỉ có chữ</p>", qtdMsgs=json.dumps([{"messageId": "1790500000001", "sender": PEER}]))
    (item,) = parse_quoted(raw)
    assert item["message_id"] == "1790500000001" and item["sender_mri"] == PEER


def test_plain_message_has_nothing_quoted():
    assert parse_quoted(TEXT_ORIGINAL) == []


# ------------------------------------------------------------------ đọc chat


def test_file_quote_shows_original_files_after_reading_it(chat):
    c, install = chat
    fake = install({CHAT: [QUOTES_FILES]}, store={(CHAT, "1790062855722"): FILES_ONLY})
    res = c.get_messages(CHAT, resolve_quotes=True)

    msg = _by_id(res, "1790580175486")
    (item,) = msg["quoted"]
    assert item["resolved"] and [f["name"] for f in item["attachments"]] == list(FILE_NAMES)
    assert item["attachments"][0]["url"].endswith("/huong-dan.md")
    assert fake.single == [(CHAT, "1790062855722")]
    # Trước đây: "Người Dùng (VF-TEST)📄 📄 📄 📄" dính liền, không id, không tên file.
    assert msg["content"] == (
        "↩️ trích tin 1790062855722 của Người Dùng (VF-TEST) · 2026-09-22 07:40 "
        "📎 huong-dan.md, schema.md, logic.md, main.md\n\nvẫn dùng 4 file này đúng không"
    )
    assert msg["quotes"] == [{"message_id": "1790062855722", "sender_mri": ME}]  # watcher: không đổi


def test_text_quote_is_filled_from_the_page_without_requests(chat):
    c, install = chat
    fake = install({CHAT: [TEXT_ORIGINAL, QUOTES_TEXT]})
    res = c.get_messages(CHAT)  # mặc định (watcher, quét song song): không đọc thêm

    msg = _by_id(res, "1790500000002")
    assert msg["quoted"][0]["resolved"] and fake.single == []
    assert msg["content"].startswith(
        '↩️ trích tin 1790500000001 của Bạn Chat (VF-TEST) · 2026-09-27 09:06: "chiều nay họp lúc 3h nhé"'
    )
    assert msg["content"].endswith("\nok")


def test_default_read_does_not_fetch_and_keeps_the_preview(chat):
    c, install = chat
    fake = install({CHAT: [QUOTES_FILES]}, store={(CHAT, "1790062855722"): FILES_ONLY})
    msg = _by_id(c.get_messages(CHAT), "1790580175486")
    assert fake.single == []
    assert msg["content"].startswith('↩️ trích tin 1790062855722 của Người Dùng (VF-TEST) · 2026-09-22 07:40: "📄 📄 📄 📄"')


def test_cross_chat_forward_and_nested_quote_read_from_the_origin_chat(chat):
    c, install = chat
    fake = install(
        {CHAT: [FORWARD]},
        store={(OTHER_CHAT, "1790418152847"): FORWARDED_ORIGINAL, (OTHER_CHAT, "1790417975375"): NESTED_ORIGINAL},
    )
    c._conv_cache = [{"id": OTHER_CHAT, "name": "Nhóm Gốc"}]
    msg = _by_id(c.get_messages(CHAT, resolve_quotes=True), "1790418269081")

    forward1, nested, forward2 = msg["quoted"]
    assert forward1["resolved"] and forward1["sender_name"] == "Người Thứ Ba (VF-TEST)"
    assert nested["resolved"] and [f["name"] for f in nested["attachments"]] == ["log.txt"]
    assert not forward2["resolved"] and forward2["error"] == "không tìm thấy (404)"
    assert (CHAT, "1790418152847") not in fake.single  # không bao giờ tìm tin chuyển tiếp ở chat hiện tại
    lines = msg["content"].splitlines()
    assert lines[0] == (
        f'↪️ chuyển tiếp tin 1790418152847 của Người Thứ Ba (VF-TEST) (chat "Nhóm Gốc" `{OTHER_CHAT}`) '
        "· 2026-09-26 10:22:"
    )
    assert lines[1] == (
        f'↩️ trích tin 1790417975375 của Bạn Chat (VF-TEST) (chat "Nhóm Gốc" `{OTHER_CHAT}`) '
        '· 2026-09-26 10:19: "tin trong nhóm gốc" 📎 log.txt'
    )
    assert "câu thứ nhất" in msg["content"] and "câu thứ hai" in msg["content"]
    assert [f["name"] for f in msg["attachments"]] == ["plan.xlsx"]  # file của chính tin chuyển tiếp


def test_forbidden_origin_chat_keeps_preview_and_is_asked_once(chat):
    c, install = chat
    locked = {**FORWARD, "properties": {
        **FORWARD["properties"],
        "originalMessageContext": {**FORWARD["properties"]["originalMessageContext"], "originalThreadId": LOCKED_CHAT},
        "originalMessageContext1": {**FORWARD["properties"]["originalMessageContext1"], "originalThreadId": LOCKED_CHAT},
    }}
    fake = install({CHAT: [locked]}, locked=(LOCKED_CHAT,))
    msg = _by_id(c.get_messages(CHAT, resolve_quotes=True), "1790418269081")

    assert len(fake.single) == 1  # một tin 403 thì cả chat đó 403
    assert all(i["error"] == "không có quyền đọc chat gốc (403)" for i in msg["quoted"])
    nested = msg["quoted"][1]
    assert quoted_line(nested, CHAT) == (
        f'↩️ trích tin 1790417975375 của Bạn Chat (VF-TEST) (chat `{LOCKED_CHAT}`) · 2026-09-26 10:19 '
        "(chưa đọc được tin gốc: không có quyền đọc chat gốc (403))"
    )
    assert "(chưa đọc" not in quoted_line(msg["quoted"][0], CHAT)  # tin chuyển tiếp đã có nguyên văn
    c.get_messages(CHAT, resolve_quotes=True)
    assert len(fake.single) == 1  # nhớ 403, không hỏi lại


def test_fetch_errors_never_break_reading(chat, monkeypatch):
    c, install = chat
    install({CHAT: [QUOTES_FILES]})

    def boom(*_a, **_k):
        raise RuntimeError("mạng rớt")

    monkeypatch.setattr(c, "_read_quoted", boom)
    msg = _by_id(c.get_messages(CHAT, resolve_quotes=True), "1790580175486")
    assert "RuntimeError" in msg["quoted"][0]["error"]
    assert msg["content"].startswith('↩️ trích tin 1790062855722 của Người Dùng (VF-TEST) · 2026-09-22 07:40: "📄 📄 📄 📄"')


def test_original_reads_are_capped_and_cached(chat, monkeypatch):
    c, install = chat
    monkeypatch.setattr(client_mod, "QUOTE_FETCH_MAX", 2)
    quoting = [
        _msg(str(1790600000000 + i), PEER, "B", _reply_html(str(1790000000000 + i), ME, "A", "x") + "<p>?</p>",
             qtdMsgs=_qtd(str(1790000000000 + i), ME))
        for i in range(4)
    ]
    store = {(CHAT, str(1790000000000 + i)): _msg(str(1790000000000 + i), ME, "A", "<p>gốc</p>") for i in range(4)}
    fake = install({CHAT: quoting}, store=store)

    res = c.get_messages(CHAT, resolve_quotes=True)
    assert len(fake.single) == 2
    errors = [m["quoted"][0]["error"] for m in res["messages"]]
    assert errors.count("") == 2 and errors.count("đã đọc đủ 2 tin gốc trong lượt này") == 2
    res = c.get_messages(CHAT, resolve_quotes=True)
    assert len(fake.single) == 4  # hai tin đã đọc lấy từ bộ nhớ đệm
    assert all(m["quoted"][0]["resolved"] for m in res["messages"])


def test_get_message_reads_an_old_message_by_id(chat):
    c, install = chat
    install({CHAT: [TEXT_ORIGINAL]}, store={(CHAT, "1790580175486"): QUOTES_FILES, (CHAT, "1790062855722"): FILES_ONLY})
    conv, msg = c.get_message(CHAT, "1790580175486")
    assert conv["id"] == CHAT and len(msg["quoted"][0]["attachments"]) == 4


# ------------------------------------------------------------------ tool


async def _call(monkeypatch, client, tool: str, args: dict, downloads: list[str] | None = None) -> str:
    from mcp.server.mcpserver import MCPServer

    import tools as tools_mod

    class FakeSharePoint:
        def download_link(self, link, target_dir=""):
            downloads.append(link)
            return f"✓ Đã tải `{link.rsplit('/', 1)[-1]}`"

    monkeypatch.setattr(tools_mod, "teams", lambda: client)
    monkeypatch.setattr(tools_mod, "sp", lambda: FakeSharePoint())
    mcp = MCPServer("t")
    tools_mod.register_all(mcp)
    res = await mcp.call_tool(tool, args)
    return res.content[0].text


@pytest.mark.anyio
async def test_read_teams_chat_lists_quoted_files_with_links(chat, monkeypatch):
    c, install = chat
    install({CHAT: [QUOTES_FILES]}, store={(CHAT, "1790062855722"): FILES_ONLY})
    text = await _call(monkeypatch, c, "read_teams_chat", {"chat_name_or_id": CHAT})
    assert "↩️ trích tin 1790062855722 của Người Dùng (VF-TEST)" in text
    assert "↩️📎 File của tin 1790062855722: [huong-dan.md](https://contoso-my.sharepoint.com/" in text
    assert text.count("](https://contoso-my") == 4


@pytest.mark.anyio
async def test_download_chat_attachments_follows_the_quote_only_when_asked(chat, monkeypatch):
    c, install = chat
    install({CHAT: [QUOTES_FILES]}, store={(CHAT, "1790062855722"): FILES_ONLY})
    downloads: list[str] = []

    text = await _call(monkeypatch, c, "download_chat_attachments",
                       {"chat_name_or_id": CHAT, "message_id": "1790580175486"}, downloads)
    assert downloads == [] and "include_quoted=true" in text and "`1790062855722`" in text

    text = await _call(monkeypatch, c, "download_chat_attachments",
                       {"chat_name_or_id": CHAT, "message_id": "1790580175486", "include_quoted": True}, downloads)
    assert sorted(d.rsplit("/", 1)[-1] for d in downloads) == sorted(FILE_NAMES)
    assert "Đã xử lý 4/4 tệp" in text and "↩️ File của tin được trích `1790062855722`:" in text


@pytest.mark.anyio
async def test_download_chat_attachments_default_scan_is_unchanged(chat, monkeypatch):
    c, install = chat
    install({CHAT: [FILES_ONLY, QUOTES_FILES]})
    downloads: list[str] = []
    text = await _call(monkeypatch, c, "download_chat_attachments", {"chat_name_or_id": CHAT, "limit": 10}, downloads)
    # Chỉ 4 file của chính tin gốc, không đếm trùng qua trích dẫn.
    assert len(downloads) == 4 and "Đã xử lý 4/4 tệp" in text


def test_download_message_images_can_take_quoted_images(chat, monkeypatch, tmp_path):
    c, install = chat
    img = ('<p><img itemtype="http://schema.skype.com/AMSImage" '
           'src="https://as-api.asm.skype.com/v1/objects/0-sa-d1-abc/views/imgo" width="250"></p>')
    original = _msg("1790500000011", PEER, "Bạn Chat (VF-TEST)", img)
    quoting = _msg("1790500000012", ME, "Người Dùng (VF-TEST)",
                   _reply_html("1790500000011", PEER, "Bạn Chat (VF-TEST)", "📷") + "<p>ảnh này</p>",
                   qtdMsgs=_qtd("1790500000011", PEER))
    install({CHAT: [quoting]}, store={(CHAT, "1790500000011"): original})

    def fake_download(url, path):
        path.write_bytes(b"PNG")
        return path

    monkeypatch.setattr(c, "download_image", fake_download)
    assert c.download_message_images(CHAT, message_id="1790500000012", target_dir=str(tmp_path)) == []
    got = c.download_message_images(CHAT, message_id="1790500000012", target_dir=str(tmp_path), include_quoted=True)
    assert len(got) == 1 and got[0]["message_id"] == "1790500000011" and got[0]["quoted_by"] == "1790500000012"
    assert got[0]["sender"] == "Bạn Chat (VF-TEST)"
