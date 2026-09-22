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
