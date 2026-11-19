"""Short-lived tokens minted from the browser's Microsoft sign-in, without Graph consent.

The configured Chromium profile already carries persistent Microsoft sign-in
cookies (``ESTSAUTH`` / ``ESTSAUTHPERSISTENT``). We replay those cookies only to
the login host and run a first-party web app's own SPA authorization (the same
silent PKCE flow the browser runs: ``prompt=none``, ``response_mode=fragment``)
to mint an access token for that app's resource. No device login, app
registration, Graph grant, refresh-token cache, or tenant-admin permission is
introduced. Only the access token is kept, in memory, until shortly before it
expires; a returned refresh token is discarded and the next token is minted
from the browser cookies again.

Subclasses say which app (client id, redirect URI, scope, expected audience)
through :meth:`SpaTokenManager._settings`: Outlook Web for mail
(``outlook.auth``), Teams Web for the chat-list sections (``teams.csa``).
"""

from __future__ import annotations

import base64
import hashlib
import secrets
import threading
import time
import urllib.parse
from dataclasses import dataclass

from common.chrome_cookies import ChromeCookieDecryptor
from common.errors import AuthExpiredError
from common.http import capture_redirect_fragment, request_json
from teams.auth import decode_jwt_claims


def _b64url(value: bytes) -> str:
    return base64.urlsafe_b64encode(value).rstrip(b"=").decode("ascii")


@dataclass(frozen=True)
class SpaSettings:
    #: Shown in errors: "Outlook Web", "Teams (mục chat)".
    product: str
    client_id: str
    tenant_id: str
    login_host: str
    #: Sent as ``Origin`` to ``/token`` (the SPA's origin).
    origin: str
    scope: str
    redirect_uri: str
    #: The token's ``aud`` must equal this.
    audience: str
    #: What the human opens in Chrome to renew the session.
    open_url: str


class SpaTokenManager:
    """Mint and cache one first-party SPA's short-lived token from browser cookies."""

    def __init__(self) -> None:
        self._token = ""
        self._expires_at = 0.0
        self._lock = threading.Lock()

    def _settings(self) -> SpaSettings:
        raise NotImplementedError

    def _username(self) -> str:
        """The account to pick in the browser session (``login_hint``)."""
        raise NotImplementedError

    def clear(self) -> None:
        with self._lock:
            self._token = ""
            self._expires_at = 0.0

    def _mint_token(self) -> tuple[str, float]:
        cfg = self._settings()
        login_host = cfg.login_host.strip()
        open_url = cfg.open_url.rstrip("/")
        cookies = ChromeCookieDecryptor.get_cookies_for_domain(login_host, use_cache=False)
        if not (cookies.get("ESTSAUTH") or cookies.get("ESTSAUTHPERSISTENT")):
            raise AuthExpiredError(
                f"Chrome không có phiên đăng nhập Microsoft bền để mở {cfg.product}.",
                f"Mở {open_url} trong Chrome, đăng nhập tài khoản công ty và chọn "
                "'Stay signed in', rồi thử lại. Không cần cấp Graph permission.",
            )

        verifier = _b64url(secrets.token_bytes(48))
        challenge = _b64url(hashlib.sha256(verifier.encode("ascii")).digest())
        state = secrets.token_urlsafe(24)
        tenant = urllib.parse.quote(cfg.tenant_id.strip() or "organizations", safe="")
        authorize_url = f"https://{login_host}/{tenant}/oauth2/v2.0/authorize?"
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
            context=f"dùng phiên Chrome đăng nhập {cfg.product}",
        )
        if not code:
            raise AuthExpiredError(
                f"Phiên Microsoft trong Chrome không thể đăng nhập ngầm vào {cfg.product}.",
                f"Mở {open_url} trong Chrome, chọn đúng tài khoản và đăng nhập lại với "
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
            f"https://{login_host}/{tenant}/oauth2/v2.0/token",
            method="POST",
            headers={
                "Accept": "application/json",
                "Content-Type": "application/x-www-form-urlencoded",
                "Origin": cfg.origin.rstrip("/"),
            },
            data=token_body,
            context=f"đổi phiên Chrome lấy token {cfg.product}",
        )
        # Only the access token is kept; a refresh token in the response is dropped here.
        token = str(response.get("access_token") or "")
        if not token:
            raise AuthExpiredError(
                f"{cfg.product} không trả access token từ phiên Chrome.",
                f"Mở {open_url} trong Chrome và đăng nhập lại.",
            )
        claims = decode_jwt_claims(token)
        if claims.get("aud") != cfg.audience:
            raise AuthExpiredError(
                f"Phiên Chrome trả token không dành cho {cfg.product}.",
                f"Mở {open_url} trong Chrome và đăng nhập đúng tài khoản công ty.",
            )
        expires_at = float(claims.get("exp") or (time.time() + int(response.get("expires_in") or 3600)))
        return token, expires_at

    def get_token(self, force_refresh: bool = False) -> str:
        """Return the in-memory token, minting a new one from Chrome cookies near expiry."""
        with self._lock:
            if not force_refresh and self._token and time.time() < self._expires_at - 60:
                return self._token
            self._token, self._expires_at = self._mint_token()
            return self._token
