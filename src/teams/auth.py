"""Teams session extraction from the local Chrome profile.

Two distinct credentials live in the browser session and are used for different
Teams back-ends:

``skypetoken_asm``
    Chat Service token (``https://{region}.ng.msg.teams.microsoft.com/v1``).
    Sent as the ``Authentication: skypetoken=<jwt>`` header.

``authtoken``
    Middle-tier bearer token (audience ``https://api.spaces.skype.com``) used by
    ``https://teams.microsoft.com/api/...`` endpoints such as calendar. Stored
    URL-encoded as ``Bearer=<jwt>&Origin=...``.
"""

from __future__ import annotations

import base64
import json
import threading
import time
import urllib.parse
from typing import Any

from common.chrome_cookies import ChromeCookieDecryptor
from common.errors import AuthExpiredError, CookieError
from common.identity import Identity

_TEAMS_COOKIE_DOMAIN = "teams.microsoft.com"
_CHAT_COOKIE_DOMAIN = "asyncgw.teams.microsoft.com"


def decode_jwt_claims(token: str) -> dict[str, Any]:
    """Decode a JWT payload without verifying the signature.

    Verification is Microsoft's job here: the token is replayed to Microsoft,
    which rejects anything invalid. We only read routing hints (region, expiry).
    """
    parts = token.split(".")
    if len(parts) < 2:
        raise AuthExpiredError(
            "Token không đúng định dạng JWT.",
            "Mở lại https://teams.microsoft.com trong Chrome để tạo phiên mới.",
        )
    payload = parts[1] + "=" * ((4 - len(parts[1]) % 4) % 4)
    try:
        return json.loads(base64.urlsafe_b64decode(payload).decode("utf-8", errors="ignore"))
    except Exception as exc:
        raise AuthExpiredError(
            f"Không giải mã được payload của token: {exc}",
            "Mở lại https://teams.microsoft.com trong Chrome.",
        ) from exc


class TeamsAuthManager:
    _cached_auth: dict[str, Any] | None = None
    _lock = threading.Lock()

    @classmethod
    def get_auth(cls, force_refresh: bool = False) -> dict[str, Any]:
        now = time.time()
        with cls._lock:
            cached = cls._cached_auth
            if not force_refresh and cached and cached.get("exp", 0) > (now + 120):
                return cached

            if force_refresh:
                ChromeCookieDecryptor.clear_cache()

            cookies = ChromeCookieDecryptor.get_cookies_for_domain(
                domain_pattern=_CHAT_COOKIE_DOMAIN,
                cookie_names=["skypetoken_asm"],
                use_cache=not force_refresh,
            )
            skypetoken = cookies.get("skypetoken_asm")
            if not skypetoken:
                raise AuthExpiredError(
                    "Không tìm thấy cookie 'skypetoken_asm' của Microsoft Teams trong Chrome.",
                    "Mở https://teams.microsoft.com trong Chrome và đăng nhập, rồi thử lại.",
                )

            claims = decode_jwt_claims(skypetoken)
            exp = float(claims.get("exp", now + 3600))
            if exp <= now:
                raise AuthExpiredError(
                    "Skypetoken của Teams đã hết hạn.",
                    "Mở lại (hoặc tải lại) https://teams.microsoft.com trong Chrome để làm mới phiên.",
                )

            region = claims.get("rgn", "apac")

            # Optional: middle-tier bearer token. Absence is not fatal - it only
            # disables calendar-style endpoints.
            middle_tier_token = ""
            middle_tier_claims: dict[str, Any] = {}
            try:
                mt_cookies = ChromeCookieDecryptor.get_cookies_for_domain(
                    domain_pattern=_TEAMS_COOKIE_DOMAIN,
                    cookie_names=["authtoken"],
                    use_cache=not force_refresh,
                )
                raw = urllib.parse.unquote(mt_cookies.get("authtoken", ""))
                if raw:
                    for chunk in raw.split("&"):
                        if chunk.startswith("Bearer="):
                            middle_tier_token = chunk[len("Bearer=") :]
                            break
                    else:
                        middle_tier_token = raw
                    if middle_tier_token:
                        middle_tier_claims = decode_jwt_claims(middle_tier_token)
            except (CookieError, AuthExpiredError):
                middle_tier_token = ""
                middle_tier_claims = {}

            merged_claims = {**middle_tier_claims, **claims}
            identity = Identity.from_claims(merged_claims)

            auth_data = {
                "token": skypetoken,
                "region": region,
                "skypeid": claims.get("skypeid", ""),
                "exp": exp,
                "claims": claims,
                "identity": identity,
                "middle_tier_token": middle_tier_token,
                "middle_tier_exp": float(middle_tier_claims.get("exp", 0) or 0),
                "base_url": f"https://{region}.ng.msg.teams.microsoft.com/v1",
            }
            cls._cached_auth = auth_data
            return auth_data

    @classmethod
    def get_identity(cls) -> Identity:
        return cls.get_auth()["identity"]

    @classmethod
    def invalidate(cls) -> None:
        with cls._lock:
            cls._cached_auth = None
