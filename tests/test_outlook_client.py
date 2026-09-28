"""Offline behavior tests for Outlook browser-session mail."""

from __future__ import annotations

import base64
import json
import time
import urllib.parse
from types import SimpleNamespace

import pytest

from common.config import reset_config_cache
from common.errors import AuthExpiredError, ConfigError
from outlook.auth import MailAuthManager
from outlook.client import OutlookMailClient, clean_mail_body


class _Auth:
    def get_token(self, force_refresh: bool = False) -> str:
        return "mail-token"


def _raw_message(**overrides):
    raw = {
        "Id": "AAMk-1+/=",
        "ConversationId": "conv-1",
        "Subject": "Review architecture",
        "From": {"EmailAddress": {"Name": "Nam Sơn", "Address": "nam@example.com"}},
        "ToRecipients": [{"EmailAddress": {"Address": "me@example.com"}}],
        "CcRecipients": [],
        "ReceivedDateTime": "2026-09-21T01:00:00Z",
        "SentDateTime": "2026-09-21T00:59:00Z",
        "IsRead": False,
        "HasAttachments": True,
        "Importance": "High",
        "Flag": {"FlagStatus": "Flagged"},
        "BodyPreview": "Please review",
        "WebLink": "https://outlook.office.com/mail/id/AAMk-1",
    }
    raw.update(overrides)
    return raw


def _jwt(**claims) -> str:
    def part(value):
        return base64.urlsafe_b64encode(json.dumps(value).encode()).rstrip(b"=").decode()

    return f"{part({'alg': 'none'})}.{part(claims)}.signature"


def test_list_messages_uses_outlook_fields_and_returns_mail_metadata(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    captured = {}

    def fake_outlook(path, **kwargs):
        captured["path"] = path
        captured.update(kwargs)
        return {"value": [_raw_message()]}

    monkeypatch.setattr(client, "_outlook_json", fake_outlook)
    messages = client.list_messages(folder="inbox", unread_only=True, since="2026-09-20T00:00:00+07:00", limit=5)

    query = urllib.parse.parse_qs(urllib.parse.urlsplit(captured["path"]).query)
    assert captured["path"].startswith("/me/mailfolders/inbox/messages?")
    assert query["$filter"] == ["ReceivedDateTime ge 2026-09-19T17:00:00Z and IsRead eq false"]
    assert query["$orderby"] == ["ReceivedDateTime desc"]
    assert captured["prefer_text"] is True
    assert "Importance" in query["$select"][0].split(",") and "Flag" in query["$select"][0].split(",")
    assert messages[0] == {
        "id": "AAMk-1+/=",
        "conversation_id": "conv-1",
        "subject": "Review architecture",
        "from": "Nam Sơn <nam@example.com>",
        "to": ["me@example.com"],
        "cc": [],
        "received": "2026-09-21T01:00:00Z",
        "sent": "2026-09-21T00:59:00Z",
        "is_read": False,
        "has_attachments": True,
        "importance": "High",
        "is_flagged": True,
        "preview": "Please review",
        "web_link": "https://outlook.office.com/mail/id/AAMk-1",
    }


def test_search_uses_outlook_search_without_odata_filter(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    captured = {}

    def fake_outlook(path, **_kwargs):
        captured["path"] = path
        return {"value": [_raw_message(Id="match")]}

    monkeypatch.setattr(client, "_outlook_json", fake_outlook)
    messages = client.list_messages(query="architecture", limit=10)

    query = urllib.parse.parse_qs(urllib.parse.urlsplit(captured["path"]).query)
    assert query["$search"] == ['"architecture"']
    assert "$filter" not in query
    assert [message["id"] for message in messages] == ["match"]


def test_search_rejects_filters_outlook_cannot_combine():
    client = OutlookMailClient(auth=_Auth())
    with pytest.raises(ConfigError, match="không kết hợp"):
        client.list_messages(query="architecture", unread_only=True)


def test_get_message_returns_readable_full_body(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    raw = _raw_message(Body={"ContentType": "html", "Content": "<p>Hello &amp; welcome</p><p>Line 2</p>"})
    monkeypatch.setattr(client, "_outlook_json", lambda *_args, **_kwargs: raw)

    message = client.get_message("AAMk-1+/=")

    assert message["body"] == "Hello & welcome\nLine 2"
    assert clean_mail_body("<div>A<br>B</div>", "HTML") == "A\nB"


def test_send_message_uses_configured_outlook_endpoint_and_payload(monkeypatch):
    monkeypatch.setenv("MCP365_MAIL_API_ROOT", "https://mail.example.test/api/v2.0")
    reset_config_cache()
    client = OutlookMailClient(auth=_Auth())
    captured = {}

    def fake_request(url, **kwargs):
        captured["url"] = url
        captured.update(kwargs)
        return 202, b"", {}

    monkeypatch.setattr("outlook.client.request", fake_request)
    result = client.send_message(
        to=[" alice@example.com "],
        cc=["manager@example.com"],
        bcc=["audit@example.com"],
        subject="Exact subject",
        body="Exact body\nSecond line",
    )

    payload = json.loads(captured["data"])
    assert captured["url"] == "https://mail.example.test/api/v2.0/me/sendmail"
    assert captured["headers"]["Authorization"] == "Bearer mail-token"
    assert payload == {
        "Message": {
            "Subject": "Exact subject",
            "Body": {"ContentType": "Text", "Content": "Exact body\nSecond line"},
            "ToRecipients": [{"EmailAddress": {"Address": "alice@example.com"}}],
            "CcRecipients": [{"EmailAddress": {"Address": "manager@example.com"}}],
            "BccRecipients": [{"EmailAddress": {"Address": "audit@example.com"}}],
        },
        "SaveToSentItems": True,
    }
    assert result["status"] == 202


def test_mail_auth_mints_once_from_browser_cookies_and_config(monkeypatch):
    monkeypatch.setenv("MCP365_MAIL_USERNAME", "me@example.com")
    monkeypatch.setenv("MCP365_MAIL_CLIENT_ID", "outlook-client")
    monkeypatch.setenv("MCP365_MAIL_TENANT_ID", "tenant-id")
    monkeypatch.setenv("MCP365_MAIL_LOGIN_HOST", "login.example.test")
    monkeypatch.setenv("MCP365_MAIL_ORIGIN", "https://mail.example.test")
    monkeypatch.setenv("MCP365_MAIL_SCOPE", "https://mail.example.test/.default openid")
    monkeypatch.setenv("MCP365_MAIL_REDIRECT_URI", "https://mail.example.test/mail/")
    reset_config_cache()
    captured = {"token_calls": 0}

    def fake_cookies(domain, use_cache):
        captured["cookie_domain"] = domain
        return {"ESTSAUTHPERSISTENT": "browser-session"}

    def fake_redirect(url, **kwargs):
        captured["authorize_url"] = url
        captured["redirect"] = kwargs
        return "authorization-code"

    token = _jwt(aud="https://mail.example.test", exp=time.time() + 3600)

    def fake_request_json(url, **kwargs):
        captured["token_calls"] += 1
        captured["token_url"] = url
        captured["token_request"] = kwargs
        return {"access_token": token, "expires_in": 3600}

    monkeypatch.setattr("outlook.auth.ChromeCookieDecryptor.get_cookies_for_domain", fake_cookies)
    monkeypatch.setattr("outlook.auth.capture_redirect_fragment", fake_redirect)
    monkeypatch.setattr("outlook.auth.request_json", fake_request_json)

    auth = MailAuthManager()
    assert auth.get_token() == token
    assert auth.get_token() == token
    authorize = urllib.parse.urlsplit(captured["authorize_url"])
    params = urllib.parse.parse_qs(authorize.query)
    assert authorize.netloc == "login.example.test"
    assert params["client_id"] == ["outlook-client"]
    assert params["login_hint"] == ["me@example.com"]
    assert params["prompt"] == ["none"]
    assert captured["redirect"]["cookies"] == {"ESTSAUTHPERSISTENT": "browser-session"}
    assert captured["redirect"]["expected"]["state"] == params["state"][0]
    assert captured["token_url"].startswith("https://login.example.test/tenant-id/")
    assert captured["token_request"]["headers"]["Origin"] == "https://mail.example.test"
    assert captured["token_calls"] == 1


def test_mail_auth_requires_persistent_browser_session(monkeypatch):
    monkeypatch.setenv("MCP365_MAIL_USERNAME", "me@example.com")
    reset_config_cache()
    monkeypatch.setattr(
        "outlook.auth.ChromeCookieDecryptor.get_cookies_for_domain",
        lambda *_args, **_kwargs: {},
    )
    monkeypatch.setattr(
        "outlook.auth.TeamsAuthManager.get_identity",
        lambda: SimpleNamespace(upn="me@example.com"),
    )

    with pytest.raises(AuthExpiredError, match="không có phiên đăng nhập Microsoft bền"):
        MailAuthManager().get_token()


# ------------------------------------------------------------------ nâng cấp mail 28/09

from common.errors import Mcp365Error  # noqa: E402
from outlook import client as mail_mod  # noqa: E402


class _Recorder:
    """Ghi lại mọi request ghi; trả lời theo (method, đuôi URL)."""

    def __init__(self, answers=None):
        self.calls = []
        self.answers = answers or {}

    def __call__(self, url, **kwargs):
        method = kwargs.get("method", "GET")
        self.calls.append((method, url, kwargs))
        for (want_method, suffix), answer in self.answers.items():
            if method == want_method and url.split("?")[0].endswith(suffix):
                if isinstance(answer, Exception):
                    raise answer
                if callable(answer):
                    answer = answer(url, kwargs)
                return answer
        return 200, b"{}", {}


def _json(status, data):
    return status, json.dumps(data).encode(), {}


def test_markdown_escapes_first_then_formats():
    html_out = mail_mod.markdown_to_html(
        "Chào **anh** *Nam* `a<b>`\n[tài liệu](https://x.test/a?b=1&c=2) <script>x</script>\n\n"
        "- một\n- hai\n\n1. ba\n2. bốn\n[bad](javascript:alert(1))"
    )
    assert "<strong>anh</strong>" in html_out and "<em>Nam</em>" in html_out
    assert "<code>a&lt;b&gt;</code>" in html_out
    assert '<a href="https://x.test/a?b=1&amp;c=2">tài liệu</a>' in html_out
    assert "&lt;script&gt;" in html_out and "<script>" not in html_out
    assert "<ul><li>một</li><li>hai</li></ul>" in html_out and "<ol><li>ba</li><li>bốn</li></ol>" in html_out
    assert 'href="javascript' not in html_out and "[bad](javascript:alert(1))" in html_out
    assert "<br>" in html_out  # xuống dòng trong đoạn


def test_plain_text_body_is_never_interpreted():
    out = mail_mod.text_to_html("a **b** <i>c</i>\nd")
    assert "**b**" in out and "&lt;i&gt;" in out and "a **b** &lt;i&gt;c&lt;/i&gt;<br>d" in out
    with pytest.raises(ConfigError, match="body_format"):
        mail_mod.render_body_html("x", "html")


def test_plan_attachments_picks_direct_session_or_link_from_local_size(monkeypatch, tmp_path):
    monkeypatch.setattr(mail_mod, "DIRECT_ATTACH_LIMIT", 10)
    monkeypatch.setattr(mail_mod, "SESSION_ATTACH_LIMIT", 100)
    files = {}
    for name, size in (("a.txt", 9), ("b.bin", 10), ("c.bin", 85), ("d.bin", 101)):
        files[name] = tmp_path / name
        files[name].write_bytes(b"x" * size)
    plan = mail_mod.plan_attachments([str(p) for p in files.values()])
    # c.bin vừa một mình nhưng làm tổng vượt 100 → link; d.bin quá giới hạn một file → link
    assert [p["method"] for p in plan] == ["direct", "session", "link", "link"]
    assert [p["method"] for p in mail_mod.plan_attachments([str(files["a.txt"])], "link")] == ["link"]
    with pytest.raises(ConfigError, match="vượt giới hạn"):
        mail_mod.plan_attachments([str(files["d.bin"])], "attach")
    with pytest.raises(ConfigError, match="attach_mode"):
        mail_mod.plan_attachments([], "zip")
    with pytest.raises(ConfigError, match="Không tìm thấy file"):
        mail_mod.plan_attachments([str(tmp_path / "missing.txt")])


def test_new_mail_with_files_goes_draft_then_attachments_then_send(monkeypatch, tmp_path):
    monkeypatch.setattr(mail_mod, "DIRECT_ATTACH_LIMIT", 10)
    monkeypatch.setattr(mail_mod, "UPLOAD_CHUNK", 4)
    small, big = tmp_path / "a.txt", tmp_path / "b.bin"
    small.write_bytes(b"hello")
    big.write_bytes(b"0123456789")
    rec = _Recorder({
        ("POST", "/me/messages"): _json(201, {"Id": "D1", "ConversationId": "C1"}),
        ("POST", "/createuploadsession"): _json(201, {"UploadUrl": "https://up.test/session?authtoken=t"}),
        ("POST", "/send"): (202, b"", {}),
    })
    monkeypatch.setattr(mail_mod, "request", rec)
    client = OutlookMailClient(auth=_Auth())
    files = mail_mod.plan_attachments([str(small), str(big)])

    result = client.compose_and_send(
        to=["a@example.com", "A@example.com"], cc=["c@example.com"], subject="S", body="**x**",
        body_format="markdown", attachments=files,
    )

    steps = [(m, u.split("/api/v2.0")[-1]) for m, u, _ in rec.calls]
    assert steps == [
        ("POST", "/me/messages"),
        ("POST", "/me/messages/D1/attachments"),
        ("POST", "/me/messages/D1/attachments/createuploadsession"),
        ("PUT", "https://up.test/session?authtoken=t"),
        ("PUT", "https://up.test/session?authtoken=t"),
        ("PUT", "https://up.test/session?authtoken=t"),
        ("POST", "/me/messages/D1/send"),
    ]
    draft = json.loads(rec.calls[0][2]["data"])
    assert draft["Body"] == {"ContentType": "HTML", "Content": mail_mod.markdown_to_html("**x**")}
    assert draft["ToRecipients"] == [{"EmailAddress": {"Address": "a@example.com"}}]  # bỏ trùng
    direct = json.loads(rec.calls[1][2]["data"])
    assert direct["Name"] == "a.txt" and base64.b64decode(direct["ContentBytes"]) == b"hello"
    puts = [kw for m, _u, kw in rec.calls if m == "PUT"]
    assert [p["headers"]["Content-Range"] for p in puts] == ["bytes 0-3/10", "bytes 4-7/10", "bytes 8-9/10"]
    assert all("Authorization" not in p["headers"] for p in puts)  # UploadUrl mang authtoken riêng
    assert b"".join(p["data"] for p in puts) == b"0123456789"
    assert result["draft_id"] == "D1" and result["conversation_id"] == "C1" and result["status"] == 202


def test_plain_new_mail_without_files_keeps_sendmail(monkeypatch):
    rec = _Recorder({("POST", "/me/sendmail"): (202, b"", {})})
    monkeypatch.setattr(mail_mod, "request", rec)
    result = OutlookMailClient(auth=_Auth()).compose_and_send(to=["a@example.com"], subject="S", body="B")
    assert [u.rsplit("/", 1)[-1] for _m, u, _kw in rec.calls] == ["sendmail"]
    assert result["mode"] == "new"


def test_failed_attachment_discards_the_draft_and_says_not_sent(monkeypatch, tmp_path):
    small = tmp_path / "a.txt"
    small.write_bytes(b"x")
    boom = Mcp365Error("HTTP 413 quá lớn")
    rec = _Recorder({
        ("POST", "/me/messages"): _json(201, {"Id": "D1"}),
        ("POST", "/attachments"): boom,
    })
    monkeypatch.setattr(mail_mod, "request", rec)
    with pytest.raises(Mcp365Error, match="CHƯA được gửi") as excinfo:
        OutlookMailClient(auth=_Auth()).compose_and_send(
            to=["a@example.com"], subject="S", body="B", attachments=mail_mod.plan_attachments([str(small)])
        )
    assert "Bản nháp dở đã được xoá" in excinfo.value.message
    assert rec.calls[-1][0] == "DELETE" and rec.calls[-1][1].endswith("/me/messages/D1")
    assert not any(u.endswith("/send") for _m, u, _kw in rec.calls)


def test_chunk_retry_treats_invalid_start_as_already_uploaded(monkeypatch):
    attempts = []

    def flaky(url, **kwargs):
        attempts.append(kwargs["headers"]["Content-Range"])
        if len(attempts) == 1:
            raise Mcp365Error("Request timed out")  # mảnh đã tới, chỉ mất phản hồi
        err = Mcp365Error('HTTP 400 {"error":{"code":"InvalidStart"}}')
        err.http_status = 400
        raise err

    monkeypatch.setattr(mail_mod, "request", flaky)
    OutlookMailClient._put_chunk("https://up.test/s", b"abcd", 0, 8, "mảnh 1")
    assert attempts == ["bytes 0-3/8", "bytes 0-3/8"]

    def invalid_first(url, **kwargs):
        err = Mcp365Error("HTTP 400 InvalidStart")
        err.http_status = 400
        raise err

    monkeypatch.setattr(mail_mod, "request", invalid_first)
    with pytest.raises(Mcp365Error, match="InvalidStart"):
        OutlookMailClient._put_chunk("https://up.test/s", b"abcd", 0, 8, "mảnh 1")


def _original(**overrides):
    fields = {
        "Id": "M1",
        "Subject": "Kế hoạch",
        "From": {"EmailAddress": {"Name": "Nam", "Address": "nam@example.com"}},
        "ToRecipients": [{"EmailAddress": {"Address": "me@example.com"}},
                         {"EmailAddress": {"Address": "hien@example.com"}}],
        "CcRecipients": [{"EmailAddress": {"Address": "ME@example.com"}},
                         {"EmailAddress": {"Address": "boss@example.com"}}],
        "Body": {"ContentType": "text", "Content": "gốc"},
    }
    fields.update(overrides)
    return _raw_message(**fields)


def _reply_client(monkeypatch, raw):
    client = OutlookMailClient(auth=_Auth())
    client._me = {"address": "me@example.com", "name": "Me", "alias": "me"}
    monkeypatch.setattr(client, "_outlook_json", lambda *_a, **_k: raw)
    return client


def test_prepare_reply_all_drops_me_and_prefixes_subject_once(monkeypatch):
    client = _reply_client(monkeypatch, _original())
    plan = client.prepare_reply("M1", "reply_all", cc=["new@example.com", "hien@example.com"])
    assert plan["to"] == ["nam@example.com", "hien@example.com"]
    assert plan["cc"] == ["boss@example.com", "new@example.com"]  # hien đã ở To
    assert plan["subject"] == "RE: Kế hoạch"

    plan = client.prepare_reply("M1", "reply")
    assert (plan["to"], plan["cc"]) == (["nam@example.com"], [])

    client = _reply_client(monkeypatch, _original(Subject="Re: Kế hoạch", ReplyTo=[{"EmailAddress": {"Address": "list@example.com"}}]))
    plan = client.prepare_reply("M1", "reply")
    assert plan["to"] == ["list@example.com"] and plan["subject"] == "Re: Kế hoạch"

    with pytest.raises(ConfigError, match="ít nhất một người nhận"):
        client.prepare_reply("M1", "forward")
    assert client.prepare_reply("M1", "forward", to=["x@example.com"])["subject"] == "FW: Re: Kế hoạch"
    with pytest.raises(ConfigError, match="mode"):
        client.prepare_reply("M1", "answer")


def test_reply_uses_create_reply_all_and_puts_text_above_the_quote(monkeypatch):
    quoted = '<html><head></head><body dir="ltr"><div id="quote">gốc</div></body></html>'
    rec = _Recorder({
        ("POST", "/createreplyall"): _json(201, {"Id": "D2", "ConversationId": "C9",
                                                 "Body": {"ContentType": "HTML", "Content": quoted}}),
        ("POST", "/send"): (202, b"", {}),
    })
    monkeypatch.setattr(mail_mod, "request", rec)
    result = OutlookMailClient(auth=_Auth()).compose_and_send(
        to=["nam@example.com"], cc=["boss@example.com"], bcc=["audit@example.com"], subject="RE: Kế hoạch",
        body="Ok <anh>", reply_to_id="M1", mode="reply_all",
    )
    steps = [(m, u.split("/api/v2.0")[-1]) for m, u, _ in rec.calls]
    assert steps == [("POST", "/me/messages/M1/createreplyall"), ("PATCH", "/me/messages/D2"),
                     ("POST", "/me/messages/D2/send")]
    patch = json.loads(rec.calls[1][2]["data"])
    content = patch["Body"]["Content"]
    assert content.startswith('<html><head></head><body dir="ltr"><div')
    assert content.index("Ok &lt;anh&gt;") < content.index('<div id="quote">gốc</div>')
    assert patch["Subject"] == "RE: Kế hoạch"
    assert patch["CcRecipients"] == [{"EmailAddress": {"Address": "boss@example.com"}}]
    assert patch["BccRecipients"] == [{"EmailAddress": {"Address": "audit@example.com"}}]
    assert result["conversation_id"] == "C9" and result["mode"] == "reply_all"


class _FakeSharePoint:
    def __init__(self):
        self.log = []

    def upload_unique(self, path, folder):
        self.log.append(("upload", folder))
        return {"name": "a (1).txt", "drive_id": "b!me", "webUrl": "https://od/a?web=1", "fileUrl": "https://od/a"}

    def create_people_link(self, drive_id, remote, upns, link_type):
        self.log.append(("people", remote, tuple(upns)))
        return "https://od/people-link"

    def create_org_link(self, drive_id, remote, link_type):
        self.log.append(("org", remote))
        return "https://od/org-link"


def test_link_files_are_shared_with_recipient_upns_found_before_upload(monkeypatch, tmp_path):
    local = tmp_path / "a.txt"
    local.write_bytes(b"x")
    sp = _FakeSharePoint()
    client = OutlookMailClient(auth=_Auth(), sharepoint=sp)
    client._me = {"address": "me@example.com", "name": "Me", "alias": "me"}
    people = {
        "nam@example.com": [{"UserPrincipalName": "nam@corp.example",
                             "ScoredEmailAddresses": [{"Address": "Nam@example.com"}]}],
        "ext@other.test": [{"UserPrincipalName": "someone@corp.example",
                            "ScoredEmailAddresses": [{"Address": "someone@example.com"}]}],
    }

    def fake_json(path, **_kwargs):
        address = urllib.parse.unquote(path.split("$search=")[1]).strip('"')
        return {"value": people.get(address, [])}

    monkeypatch.setattr(client, "_outlook_json", fake_json)
    files = mail_mod.plan_attachments([str(local)], "link")

    shared = client._share_link_files(files, ["nam@example.com", "me@example.com"], "recipients")
    assert sp.log == [("upload", "Attachments"), ("people", "Attachments/a (1).txt", ("nam@corp.example",))]
    assert shared[0]["share_link"] == "https://od/people-link" and shared[0]["name"] == "a (1).txt"

    sp.log.clear()
    with pytest.raises(Mcp365Error, match="CHƯA được gửi.*ext@other.test"):
        client._share_link_files(files, ["nam@example.com", "ext@other.test"], "recipients")
    assert sp.log == []  # chưa tải gì lên

    assert client._share_link_files(files, ["ext@other.test"], "organization")[0]["share_link"] == "https://od/org-link"
    only_me = client._share_link_files(files, ["me@example.com"], "recipients")
    assert only_me[0]["share_link"] == "https://od/a?web=1"


def test_list_messages_sender_email_filters_and_name_searches(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    paths = []
    monkeypatch.setattr(client, "_outlook_json", lambda path, **_k: paths.append(path) or {"value": []})

    client.list_messages(sender="nam@example.com", has_attachments=True, unread_only=True)
    query = urllib.parse.parse_qs(urllib.parse.urlsplit(paths[-1]).query)
    assert query["$filter"] == ["IsRead eq false and HasAttachments eq true and "
                                "From/EmailAddress/Address eq 'nam@example.com'"]

    client.list_messages(sender="Nam Sơn", query="review")
    query = urllib.parse.parse_qs(urllib.parse.urlsplit(paths[-1]).query)
    assert query["$search"] == ['"from:\\"Nam Sơn\\" review"'] and "$filter" not in query

    with pytest.raises(ConfigError, match="không kết hợp"):
        client.list_messages(sender="Nam Sơn", unread_only=True)


def test_list_messages_finds_a_folder_by_display_name(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    paths = []

    def fake(path, **_kwargs):
        paths.append(path)
        if path.startswith("/me/mailfolders/inbox/childfolders"):
            return {"value": [{"Id": "F-important", "DisplayName": "Important"}]}
        if path.startswith("/me/mailfolders?"):
            return {"value": []}
        return {"value": []}

    monkeypatch.setattr(client, "_outlook_json", fake)
    client.list_messages(folder="important")
    assert paths[-1].startswith("/me/mailfolders/F-important/messages?")
    with pytest.raises(ConfigError, match="Không tìm thấy thư mục"):
        monkeypatch.setattr(client, "_outlook_json", lambda *_a, **_k: {"value": []})
        client.list_messages(folder="Không có")


def test_download_attachments_skips_inline_never_overwrites_and_falls_back_to_content_bytes(monkeypatch, tmp_path):
    client = OutlookMailClient(auth=_Auth())
    listing = {"value": [
        {"@odata.type": "#Microsoft.OutlookServices.FileAttachment", "Id": "A1", "Name": "báo cáo.xlsx", "Size": 9},
        {"@odata.type": "#Microsoft.OutlookServices.FileAttachment", "Id": "A2", "Name": "logo.png", "Size": 3,
         "IsInline": True},
        {"@odata.type": "#Microsoft.OutlookServices.FileAttachment", "Id": "A3", "Name": "../x.txt", "Size": 3},
    ]}

    def fake_json(path, **_kwargs):
        if path.endswith("/attachments/A3"):
            return {"ContentBytes": base64.b64encode(b"xyz").decode()}
        return listing

    def fake_to_file(url, dest, **_kwargs):
        if "/A3/" in url:
            raise Mcp365Error("HTTP 404")
        Path(dest).write_bytes(b"123456789")
        return 9

    from pathlib import Path

    monkeypatch.setattr(client, "_outlook_json", fake_json)
    monkeypatch.setattr(mail_mod, "request_to_file", fake_to_file)
    (tmp_path / "báo cáo.xlsx").write_bytes("cũ".encode())

    res = client.download_attachments("M1", target_dir=str(tmp_path))
    names = sorted(Path(f["path"]).name for f in res["saved"])
    assert names == ["_x.txt", "báo cáo (1).xlsx"]  # không ghi đè, không thoát thư mục, bỏ ảnh inline
    assert (tmp_path / "báo cáo.xlsx").read_bytes() == "cũ".encode()
    assert (tmp_path / "_x.txt").read_bytes() == b"xyz"
    assert res["matched"] == 2 and not res["failures"]

    res = client.download_attachments("M1", target_dir=str(tmp_path), name_filter="logo", include_inline=True)
    assert [f["name"] for f in res["saved"]] == ["logo.png"]
