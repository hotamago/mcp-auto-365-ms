"""FedAuth minted by finishing the Microsoft sign-in bounce (rtFa alone lands on login)."""

from __future__ import annotations

import http.cookiejar
import io
import urllib.request

from sharepoint.client import SharePointClient

HOST = "tenant-my.sharepoint.com"


class _Res(io.BytesIO):
    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False


def _fake_opener(monkeypatch, form_action: str, posted: list[str]):
    def build_opener(*handlers):
        jar = next(h.cookiejar for h in handlers if isinstance(h, urllib.request.HTTPCookieProcessor))

        class Opener:
            addheaders: list = []

            def open(self, url, data=None, timeout=None):
                if data is None:  # first hop: sign-in page with the auto-submit form
                    page = f'<form method="POST" action="{form_action}"><input type="hidden" name="code" value="c&amp;1"/></form>'
                    return _Res(page.encode())
                posted.append(f"{url}|{data.decode()}")
                jar.set_cookie(
                    http.cookiejar.Cookie(
                        0, "FedAuth", "FA", None, False, HOST, True, False, "/", True, True, None, True, None, None, {}
                    )
                )
                return _Res(b"ok")

        return Opener()

    monkeypatch.setattr(urllib.request, "build_opener", build_opener)
    monkeypatch.setattr(
        "sharepoint.client.ChromeCookieDecryptor.get_cookies_for_domain",
        lambda *a, **k: {"ESTSAUTH": "E"},
    )


def test_posts_form_back_to_host_and_returns_fedauth(monkeypatch):
    posted: list[str] = []
    _fake_opener(monkeypatch, f"https://{HOST}/_forms/default.aspx", posted)
    assert SharePointClient._mint_via_login(HOST, "RT") == "FA"
    assert posted == [f"https://{HOST}/_forms/default.aspx|code=c%261"]


def test_never_posts_to_another_host(monkeypatch):
    posted: list[str] = []
    _fake_opener(monkeypatch, "https://evil.example/_forms/default.aspx", posted)
    assert SharePointClient._mint_via_login(HOST, "RT") == ""
    assert posted == []
