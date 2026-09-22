"""Timeout enforcement and retry/backoff in the shared HTTP layer."""

from __future__ import annotations

import email.message
import http.client
import io
import urllib.error

import pytest

from common import http as http_mod
from common.errors import Mcp365Error, RateLimitedError


class _Response:
    def __init__(self, body=b"{}", status=200):
        self.status, self._body, self.headers = status, body, {}

    def read(self):
        return self._body

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def _http_error(code, retry_after=None):
    msg = email.message.Message()
    if retry_after:
        msg["Retry-After"] = retry_after
    return urllib.error.HTTPError("https://example.invalid", code, "Err", msg, io.BytesIO(b"{}"))


@pytest.fixture(autouse=True)
def _no_sleep(monkeypatch):
    monkeypatch.setattr(http_mod.time, "sleep", lambda _s: None)


def test_every_request_passes_a_timeout(monkeypatch):
    """Regression guard: the old code made 17 urlopen calls with no timeout."""
    seen = {}

    def fake(req, timeout=None):
        seen["timeout"] = timeout
        return _Response()

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    http_mod.request_json("https://example.invalid")
    assert isinstance(seen["timeout"], int | float) and seen["timeout"] > 0


def test_429_is_retried_then_succeeds(monkeypatch):
    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        if calls["n"] < 3:
            raise _http_error(429, retry_after="1")
        return _Response(b'{"ok":true}')

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    assert http_mod.request_json("https://example.invalid") == {"ok": True}
    assert calls["n"] == 3


def test_retries_are_bounded(monkeypatch):
    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        raise _http_error(503)

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(RateLimitedError):
        http_mod.request_json("https://example.invalid", max_retries=2)
    assert calls["n"] == 3  # initial attempt + 2 retries


def test_403_is_not_retried(monkeypatch):
    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        raise _http_error(403)

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(Mcp365Error):
        http_mod.request_json("https://example.invalid")
    assert calls["n"] == 1


def test_timeout_surfaces_actionable_error(monkeypatch):
    def fake(req, timeout=None):
        raise TimeoutError("too slow")

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(Mcp365Error) as excinfo:
        http_mod.request_json("https://example.invalid", max_retries=0)
    assert excinfo.value.remediation


def test_login_page_instead_of_json_is_explained(monkeypatch):
    monkeypatch.setattr(
        http_mod.urllib.request, "urlopen", lambda req, timeout=None: _Response(b"<html>Sign in</html>")
    )
    with pytest.raises(Mcp365Error) as excinfo:
        http_mod.request_json("https://example.invalid")
    assert "hết hạn" in excinfo.value.remediation


def test_server_hangup_on_a_read_is_retried(monkeypatch):
    """RemoteDisconnected is not a URLError and used to crash the watcher."""
    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        if calls["n"] == 1:
            raise http.client.RemoteDisconnected("Remote end closed connection without response")
        return _Response(b'{"ok":true}')

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    assert http_mod.request_json("https://example.invalid") == {"ok": True}
    assert calls["n"] == 2


def test_server_hangup_on_a_send_is_not_resent(monkeypatch):
    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        raise ConnectionResetError("reset")

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(Mcp365Error):
        http_mod.request_json("https://example.invalid", method="POST", data=b"{}")
    assert calls["n"] == 1


# ------------------------------------------------ typed network failures


def test_connect_failure_is_typed_as_never_sent(monkeypatch):
    """urllib wraps only connect/TLS/send failures in URLError: the request never left."""
    from common.errors import ConnectError

    def fake(req, timeout=None):
        raise urllib.error.URLError(TimeoutError("_ssl.c:989: The handshake operation timed out"))

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(ConnectError) as excinfo:
        http_mod.request("https://example.invalid", method="POST", data=b"{}", max_retries=0)
    assert excinfo.value.request_sent is False


def test_tls_eof_while_reading_is_a_typed_error_not_a_crash(monkeypatch):
    """SSL: UNEXPECTED_EOF_WHILE_READING used to escape request() as a raw traceback."""
    import ssl

    from common.errors import TransportError

    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        raise ssl.SSLEOFError(8, "[SSL: UNEXPECTED_EOF_WHILE_READING] EOF occurred in violation of protocol")

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(TransportError) as excinfo:
        http_mod.request_json("https://example.invalid", max_retries=1)
    assert calls["n"] == 2  # a GET may be repeated
    assert excinfo.value.request_sent is True


def test_post_is_not_repeated_after_a_read_timeout(monkeypatch):
    """The message may already be posted; a blind retry would post it twice."""
    from common.errors import TransportError

    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        raise TimeoutError("timed out")

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    with pytest.raises(TransportError) as excinfo:
        http_mod.request("https://example.invalid", method="POST", data=b"{}", max_retries=3)
    assert calls["n"] == 1
    assert "kiểm tra" in excinfo.value.remediation


def test_idempotent_put_is_repeated_after_a_read_timeout(monkeypatch):
    calls = {"n": 0}

    def fake(req, timeout=None):
        calls["n"] += 1
        if calls["n"] == 1:
            raise TimeoutError("timed out")
        return _Response(b"{}")

    monkeypatch.setattr(http_mod.urllib.request, "urlopen", fake)
    http_mod.request("https://example.invalid", method="PUT", data=b"x")
    assert calls["n"] == 2


def test_http_errors_carry_status_and_retry_after(monkeypatch):
    monkeypatch.setattr(
        http_mod.urllib.request, "urlopen", lambda req, timeout=None: (_ for _ in ()).throw(_http_error(503, "7"))
    )
    with pytest.raises(RateLimitedError) as excinfo:
        http_mod.request("https://example.invalid", max_retries=0)
    assert excinfo.value.http_status == 503
    assert excinfo.value.retry_after == "7"
