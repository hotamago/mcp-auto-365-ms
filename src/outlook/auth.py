"""Outlook browser-session authentication without Graph consent.

The configured Chromium profile already carries persistent Microsoft sign-in
cookies. We replay those cookies only to ``login.microsoftonline.com`` and use
Outlook Web's own first-party SPA authorization (the same flow the browser
runs) to mint an ``https://outlook.office.com`` token. No device login, app
registration, Graph grant, refresh-token cache, or tenant-admin permission is
introduced. The short-lived access token stays in memory.
"""

from __future__ import annotations

import base64
import hashlib
import secrets
import threading
import time
import urllib.parse

from common.chrome_cookies import ChromeCookieDecryptor
from common.config import get_config
from common.errors import AuthExpiredError, Mcp365Error
from common.http import capture_redirect_fragment, request_json
from teams.auth import TeamsAuthManager, decode_jwt_claims


def _b64url(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).rstrip(b"=").decode("ascii")


class MailAuthManager:
    """Mint and cache Outlook Web's short-lived token from browser cookies."""

    def __init__(self) -> None:
        self._token = ""
        self._expires_at = 0.0
        self._lock = threading.Lock()

    def clear(self) -> None:
        with self._lock:
            self._token = ""
            self._expires_at = 0.0

    def _username(self) -> str:
        configured = get_config().mail.username.strip()
        if configured:
            return configured
        try:
            username = TeamsAuthManager.get_identity().upn
        except Mcp365Error as exc:
            raise AuthExpiredError(
                "Không xác định được tài khoản Outlook để chọn trong phiên đăng nhập Chrome.",
                "Đặt MCP365_MAIL_USERNAME (ví dụ user@company.com), hoặc mở lại Teams trong Chrome để "
                "server tự suy ra tài khoản.",
            ) from exc
        if not username:
            raise AuthExpiredError(
                "Phiên Teams không chứa UPN để chọn tài khoản Outlook.",
                "Đặt MCP365_MAIL_USERNAME thành địa chỉ mail công ty.",
            )
        return username

    def _mint_token(self) -> tuple[str, float]:
        cfg = get_config().mail
        login_host = cfg.login_host.strip()
        origin = cfg.origin.rstrip("/")
        cookies = ChromeCookieDecryptor.get_cookies_for_domain(login_host, use_cache=False)
        if not (cookies.get("ESTSAUTH") or cookies.get("ESTSAUTHPERSISTENT")):
            raise AuthExpiredError(
                "Chrome không có phiên đăng nhập Microsoft bền để mở Outlook.",
                f"Mở {origin} trong Chrome, đăng nhập tài khoản công ty và chọn "
                "'Stay signed in', rồi thử lại. Không cần cấp Graph permission.",
            )

        verifier = _b64url(secrets.token_bytes(48))
        challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
        state = secrets.token_urlsafe(24)
        tenant = cfg.tenant_id.strip() or "organizations"
        authorize_url = f"https://{login_host}/{urllib.parse.quote(tenant, safe='')}/oauth2/v2.0/authorize?"
        authorize_url += urllib.parse.urlencode(
            {
                "client_id": cfg.client_id,
                "scope": cfg.scope,
                "redirect_uri": cfg.redirect_uri,
                "response_mode": "fragment",
                "response_type": "code",
                "code_challenge": challenge,
                "code_challenge_method": "S256",
                "prompt": "none",
                "login_hint": self._username(),
                "state": state,
                "nonce": secrets.token_urlsafe(24),
            }
        )
        code = capture_redirect_fragment(
            authorize_url,
            cookies=cookies,
            cookie_domain=login_host,
            key="code",
            expected={"state": state},
            context="dùng phiên Chrome đăng nhập Outlook Web",
        )
        if not code:
            raise AuthExpiredError(
                "Phiên Microsoft trong Chrome không thể đăng nhập ngầm vào Outlook Web.",
                f"Mở {origin} trong Chrome, chọn đúng tài khoản và đăng nhập lại với "
                "'Stay signed in'. Không cần chạy device login hay xin Graph permission.",
            )

        token_body = urllib.parse.urlencode(
            {
                "client_id": cfg.client_id,
                "scope": cfg.scope,
                "redirect_uri": cfg.redirect_uri,
                "grant_type": "authorization_code",
                "code": code,
                "code_verifier": verifier,
            }
        ).encode()
        response = request_json(
            f"https://{login_host}/{urllib.parse.quote(tenant, safe='')}/oauth2/v2.0/token",
            method="POST",
            headers={
                "Accept": "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
                "Origin": origin,
            },
            data=token_body,
            context="đổi phiên Chrome lấy token Outlook Web",
        )
        token = str(response.get("access_token") or "")
        if not token:
            raise AuthExpiredError(
                "Outlook Web không trả access token từ phiên Chrome.",
                f"Mở {origin} trong Chrome và đăng nhập lại.",
            )
        claims = decode_jwt_claims(token)
        if claims.get("aud") != origin:
            raise AuthExpiredError(
                "Phiên Chrome trả token không dành cho Outlook Web.",
                f"Mở {origin} trong Chrome và đăng nhập đúng tài khoản công ty.",
            )
        expires_at = float(claims.get("exp") or (time.time() + int(response.get("expires_in") or 3600)))
        return token, expires_at

    def get_token(self, force_refresh: bool = False) -> str:
        """Return Outlook Web's in-memory token, refreshing from Chrome cookies."""
        with self._lock:
            if not force_refresh and self._token and time.time() < self._expires_at - 60:
                return self._token
            self._token, self._expires_at = self._mint_token()
            return self._token
