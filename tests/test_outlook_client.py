"""Offline behavior tests for Outlook mail auth and Graph calls."""

from __future__ import annotations

import json
import urllib.parse

import pytest

from common.errors import AuthExpiredError, ConfigError
from outlook.auth import MailAuthManager
from outlook.client import OutlookMailClient, clean_mail_body


class _Auth:
    def get_token(self) -> str:
        return "mail-token"


def _raw_message(**overrides):
    raw = {
        "id": "AAMk-1+/=",
        "conversationId": "conv-1",
        "subject": "Review architecture",
        "from": {"emailAddress": {"name": "Nam Sơn", "address": "nam@example.com"}},
        "toRecipients": [{"emailAddress": {"address": "me@example.com"}}],
        "ccRecipients": [],
        "receivedDateTime": "2026-09-21T01:00:00Z",
        "sentDateTime": "2026-09-21T00:59:00Z",
        "isRead": False,
        "hasAttachments": True,
        "bodyPreview": "Please review",
        "webLink": "https://outlook.office.com/mail/id/AAMk-1",
    }
    raw.update(overrides)
    return raw


def test_list_messages_uses_folder_filters_and_returns_mail_metadata(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    captured = {}

    def fake_graph(path, **kwargs):
        captured["path"] = path
        captured.update(kwargs)
        return {"value": [_raw_message()]}

    monkeypatch.setattr(client, "_graph_json", fake_graph)
    messages = client.list_messages(
        folder="inbox", unread_only=True, since="2026-09-20T00:00:00+07:00", limit=5
    )

    query = urllib.parse.parse_qs(urllib.parse.urlsplit(captured["path"]).query)
    assert captured["path"].startswith("/me/mailFolders/inbox/messages?")
    assert query["$filter"] == ["receivedDateTime ge 2026-09-19T17:00:00Z and isRead eq false"]
    assert query["$orderby"] == ["receivedDateTime desc"]
    assert captured["prefer_text"] is True
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
        "preview": "Please review",
        "web_link": "https://outlook.office.com/mail/id/AAMk-1",
    }


def test_search_uses_graph_search_without_odata_filter(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    captured = {}

    def fake_graph(path, **_kwargs):
        captured["path"] = path
        return {"value": [_raw_message(id="match")]}

    monkeypatch.setattr(client, "_graph_json", fake_graph)
    messages = client.list_messages(query="architecture", limit=10)

    query = urllib.parse.parse_qs(urllib.parse.urlsplit(captured["path"]).query)
    assert query["$search"] == ['"architecture"']
    assert "$filter" not in query
    assert [message["id"] for message in messages] == ["match"]


def test_search_rejects_filters_graph_cannot_combine():
    client = OutlookMailClient(auth=_Auth())
    with pytest.raises(ConfigError, match="không kết hợp"):
        client.list_messages(query="architecture", unread_only=True)


def test_get_message_returns_readable_full_body(monkeypatch):
    client = OutlookMailClient(auth=_Auth())
    raw = _raw_message(body={"contentType": "html", "content": "<p>Hello &amp; welcome</p><p>Line 2</p>"})
    monkeypatch.setattr(client, "_graph_json", lambda *_args, **_kwargs: raw)

    message = client.get_message("AAMk-1+/=")

    assert message["body"] == "Hello & welcome\nLine 2"
    assert clean_mail_body("<div>A<br>B</div>", "HTML") == "A\nB"


def test_send_message_submits_exact_plain_text_and_recipient_sets(monkeypatch):
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
    assert captured["url"].endswith("/me/sendMail")
    assert captured["headers"]["Authorization"] == "Bearer mail-token"
    assert payload == {
        "message": {
            "subject": "Exact subject",
            "body": {"contentType": "Text", "content": "Exact body\nSecond line"},
            "toRecipients": [{"emailAddress": {"address": "alice@example.com"}}],
            "ccRecipients": [{"emailAddress": {"address": "manager@example.com"}}],
            "bccRecipients": [{"emailAddress": {"address": "audit@example.com"}}],
        },
        "saveToSentItems": True,
    }
    assert result["status"] == 202


def test_mail_auth_never_starts_interactive_login_implicitly(tmp_path, monkeypatch):
    class FakeApp:
        initiated = False

        def get_accounts(self):
            return []

        def initiate_device_flow(self, scopes):
            self.initiated = True
            return {"user_code": "CODE"}

    app = FakeApp()
    auth = MailAuthManager(cache_path=tmp_path / "cache.json")
    monkeypatch.setattr(auth, "_application", lambda: app)

    with pytest.raises(AuthExpiredError, match="chưa có phiên"):
        auth.get_token()
    assert app.initiated is False


def test_device_login_reports_code_then_transitions_to_connected(tmp_path, monkeypatch):
    class FakeApp:
        def get_accounts(self):
            return []

        def initiate_device_flow(self, scopes):
            assert scopes == ["Mail.Read", "Mail.Send"]
            return {
                "user_code": "ABCD-EFGH",
                "verification_uri": "https://microsoft.com/devicelogin",
                "expires_in": 900,
                "message": "Sign in",
            }

        def acquire_token_by_device_flow(self, flow):
            assert flow["user_code"] == "ABCD-EFGH"
            return {
                "access_token": "secret-not-rendered",
                "id_token_claims": {"preferred_username": "me@example.com"},
            }

    auth = MailAuthManager(cache_path=tmp_path / "cache.json")
    monkeypatch.setattr(auth, "_application", lambda: FakeApp())

    started = auth.start_device_login()
    assert started["status"] == "pending"
    assert started["user_code"] == "ABCD-EFGH"
    auth._thread.join(timeout=1)
    assert auth.login_status() == {"status": "connected", "username": "me@example.com"}
