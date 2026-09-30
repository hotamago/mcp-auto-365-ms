"""Một-hoặc-nhiều (common.multi) và các tool nhận nhiều file / nhiều tin / nhiều chat (30/09).

Không gọi mạng: client Teams/SharePoint/Outlook đều là đồ giả.
"""

from __future__ import annotations

import threading
import time

import pytest
from mcp.server.mcpserver import MCPServer
from mcp.server.mcpserver.exceptions import ToolError

import tools as tools_mod
from common import multi
from common.errors import Mcp365Error

# ------------------------------------------------------------------ helpers


def test_as_list_takes_one_many_or_a_json_list_and_keeps_order():
    assert multi.as_list(None) == []
    assert multi.as_list("", [], None) == []
    assert multi.as_list("a.pdf") == ["a.pdf"]
    assert multi.as_list(" a.pdf ", ["b.pdf", "a.pdf", ""], "c.pdf") == ["a.pdf", "b.pdf", "c.pdf"]
    assert multi.as_list('["x.pdf", "y.pdf"]') == ["x.pdf", "y.pdf"]  # harness gửi mảng thành chuỗi
    assert multi.as_list("[draft] notes.md") == ["[draft] notes.md"]  # không phải JSON: một đường dẫn


def test_local_files_resolves_dedupes_and_lists_every_missing_file(tmp_path):
    a, b = tmp_path / "a.pdf", tmp_path / "b.pdf"
    a.write_bytes(b"1")
    b.write_bytes(b"2")
    assert multi.local_files(str(a), [str(b), str(tmp_path / "." / "a.pdf")]) == [a, b]
    with pytest.raises(Mcp365Error) as one:
        multi.local_files(str(a), [str(tmp_path / "x.pdf")])
    assert one.value.message == f"Không tìm thấy file đính kèm: {tmp_path / 'x.pdf'}"
    with pytest.raises(Mcp365Error) as many:
        multi.local_files([str(tmp_path / "x.pdf"), str(a), str(tmp_path / "y.pdf")])
    assert "Không tìm thấy 2 file đính kèm" in many.value.message
    assert "x.pdf" in many.value.message and "y.pdf" in many.value.message


def test_same_names_ignores_case_like_sharepoint(tmp_path):
    assert multi.same_names([tmp_path / "a" / "Spec.pdf", tmp_path / "b" / "spec.PDF", tmp_path / "c.pdf"]) == [
        "spec.PDF"
    ]
    assert multi.same_names([tmp_path / "a.pdf", tmp_path / "b.pdf"]) == []


def test_run_ordered_runs_in_parallel_and_returns_in_input_order():
    barrier = threading.Barrier(3, timeout=5)  # chỉ qua được khi 3 việc chạy cùng lúc

    def work(n: int) -> int:
        barrier.wait()
        time.sleep(0.05 * (3 - n))  # việc đầu xong sau cùng
        return n * 10

    assert multi.run_ordered(work, [0, 1, 2], max_workers=3) == [(0, None), (10, None), (20, None)]


def test_run_ordered_keeps_errors_per_item_and_the_rest_still_run():
    def work(n: int) -> int:
        if n == 1:
            raise Mcp365Error("hỏng")
        return n

    out = multi.run_ordered(work, [0, 1, 2], max_workers=3)
    assert out[0] == (0, None) and out[2] == (2, None)
    assert out[1][0] is None and isinstance(out[1][1], Mcp365Error)


def test_run_ordered_runs_same_key_one_after_another_in_order():
    active, peak, seen = [0], [0], []
    lock = threading.Lock()

    def work(item: tuple[str, int]) -> None:
        with lock:
            active[0] += 1
            peak[0] = max(peak[0], active[0])
        time.sleep(0.02)
        with lock:
            seen.append(item)
            active[0] -= 1

    items = [("a.pdf", 1), ("A.PDF", 2), ("a.pdf", 3)]
    multi.run_ordered(work, items, key=lambda i: i[0].casefold(), max_workers=4)
    assert peak[0] == 1 and seen == items


def test_run_ordered_carries_the_tool_timeout_into_worker_threads():
    from common.http import timeout_override, timeout_scope

    with timeout_scope(42):
        out = multi.run_ordered(lambda _n: timeout_override(), [1, 2, 3], max_workers=3)
    assert [v for v, _e in out] == [42.0, 42.0, 42.0]


def test_run_ordered_single_item_or_one_worker_stays_on_the_calling_thread(monkeypatch):
    me = threading.get_ident()
    assert multi.run_ordered(lambda _n: threading.get_ident(), [1]) == [(me, None)]
    monkeypatch.setenv("MCP365_TRANSFER_WORKERS", "1")
    from common.config import reset_config_cache

    reset_config_cache()
    assert multi.transfer_workers() == 1
    assert [v for v, _e in multi.run_ordered(lambda _n: threading.get_ident(), [1, 2, 3])] == [me, me, me]


# ------------------------------------------------------------ tool fixtures


@pytest.fixture
def mcp():
    server = MCPServer("t")
    tools_mod.register_all(server)
    return server


async def _refusal(mcp, tool: str, args: dict) -> str:
    with pytest.raises(ToolError) as excinfo:
        await mcp.call_tool(tool, args)
    return str(excinfo.value)


def _files(tmp_path, *names: str, size: int = 3):
    out = []
    for name in names:
        path = tmp_path / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(b"x" * size)
        out.append(path)
    return out


class _FakeTeams:
    def __init__(self, conv=None):
        self.conv = conv or {"id": "19:dev@thread.v2", "name": "Dev team", "type": "GroupChat"}
        self.sent: list[dict] = []

    def find_conversation(self, _identifier):
        return self.conv

    def resolve_mentions(self, *_a, **_k):
        return []

    def send_message(self, **kwargs):
        self.sent.append(kwargs)
        paths = kwargs.get("file_paths") or ([kwargs["file_path"]] if kwargs.get("file_path") else [])
        files = [{"name": p.rsplit("/", 1)[-1], "share_link": f"https://od/{i}", "webUrl": "w", "location": "OneDrive"}
                 for i, p in enumerate(paths)]
        return {"conversation_id": self.conv["id"], "conversation_name": self.conv["name"], "message_id": "9",
                "message_sent": kwargs["message"], "reply_to_id": None, "mentioned": [],
                "attached_file": files[0] if files else None, "attached_files": files}


# ------------------------------------------------------------ send_teams_message


@pytest.mark.anyio
async def test_teams_draft_lists_every_file_with_size_before_anything_happens(mcp, monkeypatch, tmp_path):
    fake = _FakeTeams()
    monkeypatch.setattr(tools_mod, "teams", lambda: fake)
    a, b = _files(tmp_path, "a.pdf", "b.pdf", size=2048)

    text = await _refusal(mcp, "send_teams_message", {
        "chat_name_or_id": "Dev team", "message": "hai file", "is_user_confirm": False,
        "file_path": str(a), "file_paths": [str(b)],
    })
    assert "CHƯA GỬI" in text and fake.sent == []
    assert "**Đính kèm (2 file, cùng một tin)**" in text
    assert text.index(f"1. `{a}` · 2.0 KB") < text.index(f"2. `{b}` · 2.0 KB")
    assert "chỉ các thành viên của chat này" in text


@pytest.mark.anyio
async def test_teams_multi_files_missing_one_is_refused_before_the_draft(mcp, monkeypatch, tmp_path):
    fake = _FakeTeams()
    monkeypatch.setattr(tools_mod, "teams", lambda: fake)
    (a,) = _files(tmp_path, "a.pdf")
    text = await _refusal(mcp, "send_teams_message", {
        "chat_name_or_id": "Dev team", "message": "x", "is_user_confirm": True,
        "file_paths": [str(a), str(tmp_path / "missing.pdf")],
    })
    assert "Không tìm thấy file đính kèm" in text and "missing.pdf" in text and fake.sent == []


@pytest.mark.anyio
async def test_teams_channel_refuses_two_files_with_the_same_name(mcp, monkeypatch, tmp_path):
    fake = _FakeTeams({"id": "19:c@thread.tacv2", "name": "[VF] #General", "type": "Channel"})
    monkeypatch.setattr(tools_mod, "teams", lambda: fake)
    a, b = _files(tmp_path, "x/spec.pdf", "y/spec.pdf")
    text = await _refusal(mcp, "send_teams_message", {
        "chat_name_or_id": "General", "message": "x", "is_user_confirm": False, "file_paths": [str(a), str(b)],
    })
    assert "trùng nhau (spec.pdf)" in text and "CHƯA GỬI" not in text and fake.sent == []


@pytest.mark.anyio
async def test_teams_approved_send_passes_every_file_in_order(mcp, monkeypatch, tmp_path):
    fake = _FakeTeams()
    monkeypatch.setattr(tools_mod, "teams", lambda: fake)
    a, b, c = _files(tmp_path, "a.pdf", "b.pdf", "c.pdf")
    res = await mcp.call_tool("send_teams_message", {
        "chat_name_or_id": "Dev team", "message": "ba file", "is_user_confirm": True,
        "file_path": str(b), "file_paths": f'["{a}", "{c}", "{b}"]',  # chuỗi JSON cũng được
    })
    assert fake.sent[0]["file_paths"] == [str(b), str(a), str(c)] and "file_path" not in fake.sent[0]
    text = res.content[0].text
    assert "**File đính kèm (3):**" in text and text.index("[b.pdf]") < text.index("[a.pdf]") < text.index("[c.pdf]")


@pytest.mark.anyio
async def test_teams_single_file_call_is_unchanged(mcp, monkeypatch):
    fake = _FakeTeams()
    monkeypatch.setattr(tools_mod, "teams", lambda: fake)
    text = await _refusal(mcp, "send_teams_message", {
        "chat_name_or_id": "Dev team", "message": "đây", "file_path": "/tmp/a.xlsx", "is_user_confirm": False,
    })
    assert "**Đính kèm:** `/tmp/a.xlsx` → OneDrive của bạn › `Microsoft Teams Chat Files`" in text
    await mcp.call_tool("send_teams_message", {
        "chat_name_or_id": "Dev team", "message": "đây", "file_path": "/tmp/a.xlsx", "is_user_confirm": True,
    })
    assert fake.sent[0]["file_path"] == "/tmp/a.xlsx" and "file_paths" not in fake.sent[0]
    await mcp.call_tool("send_teams_message", {"chat_name_or_id": "Dev team", "message": "chữ", "is_user_confirm": True})
    assert fake.sent[1]["file_path"] is None


# ------------------------------------------------------------ SharePoint / OneDrive uploads


class _FakeSP:
    def __init__(self, fail: set[str] | None = None):
        self.fail = fail or set()
        self.calls: list[tuple] = []

    def upload_file(self, path, folder, name=None):
        self.calls.append(("upload", path, folder, name))
        if path.rsplit("/", 1)[-1] in self.fail:
            raise Mcp365Error(f"HTTP 507 khi tải {path.rsplit('/', 1)[-1]}", "Dọn chỗ trên site.")
        n = path.rsplit("/", 1)[-1]
        return {"name": n, "size": 3, "id": f"id-{n}", "webUrl": f"https://sp/{n}?web=1", "fileUrl": f"https://sp/{n}",
                "folder": folder}

    def share_file_onedrive(self, path, folder, link_type, name):
        res = self.upload_file(path, folder, name or None)
        return {**res, "share_link": f"https://od/share/{res['name']}", "link_type": link_type}


@pytest.mark.anyio
async def test_upload_many_files_needs_one_approval_of_the_whole_list(mcp, monkeypatch, tmp_path):
    fake = _FakeSP(fail={"b.md"})
    monkeypatch.setattr(tools_mod, "sp", lambda: fake)
    a, b, c = _files(tmp_path, "a.md", "b.md", "c.md")
    args = {"target_folder_url_or_path": "Docs/X", "local_file_paths": [str(a), str(b), str(c)]}

    text = await _refusal(mcp, "upload_sharepoint_file", {**args, "is_user_confirm": False})
    assert "Tải 3 file lên SharePoint" in text and fake.calls == []
    assert text.index(f"1. `{a}` → `a.md`") < text.index(f"2. `{b}` → `b.md`") < text.index(f"3. `{c}`")

    out = (await mcp.call_tool("upload_sharepoint_file", {**args, "is_user_confirm": True})).content[0].text
    assert "Đã tải 2/3 file" in out and out.index("`a.md`") < out.index("`c.md`")
    assert "Thất bại (1/3)" in out and "HTTP 507 khi tải b.md" in out
    assert sorted(c[1] for c in fake.calls) == sorted([str(a), str(b), str(c)])

    fake.fail = {"a.md", "b.md", "c.md"}
    text = await _refusal(mcp, "upload_sharepoint_file", {**args, "is_user_confirm": True})
    assert "Không tải được file nào" in text and "Dọn chỗ trên site." in text


@pytest.mark.anyio
async def test_upload_many_refuses_same_names_and_a_rename(mcp, monkeypatch, tmp_path):
    monkeypatch.setattr(tools_mod, "sp", lambda: _FakeSP())
    a, b = _files(tmp_path, "x/a.md", "y/a.md")
    text = await _refusal(mcp, "upload_sharepoint_file", {
        "target_folder_url_or_path": "Docs", "local_file_paths": [str(a), str(b)], "is_user_confirm": False})
    assert "Tên file trùng nhau (a.md)" in text
    text = await _refusal(mcp, "upload_sharepoint_file", {
        "target_folder_url_or_path": "Docs", "local_file_paths": [str(a)], "local_file_path": str(tmp_path / "x" / "a.md"),
        "target_file_name": "b.md", "is_user_confirm": False})
    assert "chỉ dùng khi tải một file" in text


@pytest.mark.anyio
async def test_upload_single_file_call_is_unchanged(mcp, monkeypatch, tmp_path):
    fake = _FakeSP()
    monkeypatch.setattr(tools_mod, "sp", lambda: fake)
    text = await _refusal(mcp, "upload_sharepoint_file", {
        "local_file_path": "/tmp/r.md", "target_folder_url_or_path": "Docs", "target_file_name": "s.md",
        "is_user_confirm": False})
    assert "`/tmp/r.md` → `s.md`" in text
    out = (await mcp.call_tool("upload_sharepoint_file", {
        "local_file_path": "/tmp/r.md", "target_folder_url_or_path": "Docs", "target_file_name": "s.md",
        "is_user_confirm": True})).content[0].text
    assert fake.calls == [("upload", "/tmp/r.md", "Docs", "s.md")] and out.startswith("✓ Đã tải `r.md`")


@pytest.mark.anyio
async def test_share_many_files_on_onedrive_gives_one_link_each(mcp, monkeypatch, tmp_path):
    fake = _FakeSP()
    monkeypatch.setattr(tools_mod, "sp", lambda: fake)
    a, b = _files(tmp_path, "a.pdf", "b.pdf")
    text = await _refusal(mcp, "share_file_onedrive", {"local_file_paths": [str(a), str(b)], "is_user_confirm": False})
    assert "Tải 2 file lên OneDrive" in text and "mỗi file một link" in text and fake.calls == []
    out = (await mcp.call_tool("share_file_onedrive", {"local_file_paths": [str(a), str(b)],
                                                       "is_user_confirm": True})).content[0].text
    assert "https://od/share/a.pdf" in out and "https://od/share/b.pdf" in out and "Đã tải 2/2 file" in out


# ------------------------------------------------------------ mail


@pytest.mark.anyio
async def test_mail_attachments_take_a_single_path_string(mcp, monkeypatch, tmp_path):
    monkeypatch.setattr(tools_mod, "outlook", lambda: pytest.fail("nothing before approval"))
    a, b = _files(tmp_path, "a.txt", "b.txt")
    base = {"to": ["me@example.com"], "subject": "S", "body": "B", "is_user_confirm": False}
    text = await _refusal(mcp, "send_email", {**base, "attachments": str(a)})
    assert "`a.txt`" in text
    text = await _refusal(mcp, "send_email", {**base, "attachments": f'["{a}", "{b}"]'})
    assert text.index("`a.txt`") < text.index("`b.txt`")
    text = await _refusal(mcp, "send_email", {**base, "attachments": [str(a), str(tmp_path / "x"), str(tmp_path / "y")]})
    assert "Không tìm thấy 2 file đính kèm" in text and "CHƯA GỬI" not in text


@pytest.mark.anyio
async def test_download_attachments_of_several_emails(mcp, monkeypatch):
    class FakeMail:
        def download_attachments(self, mid, **_kw):
            if mid == "bad":
                raise Mcp365Error("HTTP 404 Not Found")
            return {"target_dir": "/d", "matched": 1, "saved": [{"name": f"{mid}.pdf", "path": f"/d/{mid}.pdf",
                                                                 "size": 5}], "failures": []}

    monkeypatch.setattr(tools_mod, "outlook", lambda: FakeMail())
    out = (await mcp.call_tool("download_email_attachments",
                               {"message_id": "m1", "message_ids": ["m2", "bad"]})).content[0].text
    assert out.index("## Mail `m1`") < out.index("## Mail `m2`") and "/d/m2.pdf" in out
    assert "Không đọc được 1/3 mail" in out and "`bad`: HTTP 404" in out
    one = (await mcp.call_tool("download_email_attachments", {"message_id": "m1"})).content[0].text
    assert one.startswith("# Đã tải 1/1 file đính kèm")


# ------------------------------------------------------------ read / download Teams


class _FakeChats:
    def __init__(self):
        self.calls: list[str] = []

    def get_messages(self, chat, **_kw):
        self.calls.append(chat)
        if chat == "hỏng":
            raise Mcp365Error("HTTP 403 Forbidden")
        time.sleep(0.05 if chat == "A" else 0)
        return {"conversation_id": f"19:{chat}", "conversation_name": f"Chat {chat}", "conversation_type": "GroupChat",
                "messages": [{"timestamp": "2026-09-30 09:00", "sender": "X", "content": f"tin trong {chat}"}]}


@pytest.mark.anyio
async def test_read_several_chats_in_one_call(mcp, monkeypatch, tmp_path):
    fake = _FakeChats()
    monkeypatch.setattr(tools_mod, "teams", lambda: fake)
    out_file = tmp_path / "t.md"
    out = (await mcp.call_tool("read_teams_chat", {"chat_name_or_id": "A", "chat_names_or_ids": ["B", "hỏng"],
                                                   "output_file": str(out_file)})).content[0].text
    assert out.index("# Chat: Chat A") < out.index("# Chat: Chat B")  # A chậm hơn vẫn đứng trước
    assert "Không đọc được 1/3 chat" in out and "hỏng: HTTP 403" in out
    assert "tin trong B" in out_file.read_text(encoding="utf-8")
    single = (await mcp.call_tool("read_teams_chat", {"chat_name_or_id": "B"})).content[0].text
    assert single.startswith("# Chat: Chat B") and "---\n\n# Chat" not in single


@pytest.mark.anyio
async def test_download_attachments_of_several_teams_messages(mcp, monkeypatch):
    def msg(mid, name):
        return {"id": mid, "attachments": [{"name": name, "url": f"https://t-my.sharepoint.com/{name}"}],
                "sharepoint_links": [], "images": []}

    class FakeTeams:
        def get_messages(self, chat, limit=50, **_kw):
            return {"conversation_id": "19:c", "conversation_name": "Dev", "messages": [msg("1", "a.pdf"), msg("2", "b.pdf")]}

        def get_message(self, chat, mid, scan=50, resolve_quotes=False):
            if mid == "3":
                return {"id": "19:c", "name": "Dev"}, msg("3", "c.pdf")
            raise Mcp365Error(f"Không tìm thấy message ID '{mid}'")

    downloads: list[str] = []

    class FakeSP:
        def download_link(self, link, target_dir=""):
            time.sleep(0.03 if link.endswith("b.pdf") else 0)
            downloads.append(link)
            return f"✓ {link.rsplit('/', 1)[-1]}"

    monkeypatch.setattr(tools_mod, "teams", lambda: FakeTeams())
    monkeypatch.setattr(tools_mod, "sp", lambda: FakeSP())
    out = (await mcp.call_tool("download_chat_attachments", {
        "chat_name_or_id": "Dev", "message_ids": ["2", "3", "9", "1"], "limit": 10})).content[0].text
    assert out.index("✓ b.pdf") < out.index("✓ c.pdf") < out.index("✓ a.pdf")  # thứ tự id đưa vào
    assert "Đã xử lý 3/3 tệp" in out and "Không đọc được 1/4 tin" in out and "'9'" in out
    out = (await mcp.call_tool("download_chat_attachments", {
        "chat_name_or_id": "Dev", "message_ids": ["2", "3", "1"], "limit": 2})).content[0].text
    assert "Còn 1 tệp nữa chưa tải" in out
