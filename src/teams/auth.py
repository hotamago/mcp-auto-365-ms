"""Authentication module for Microsoft Teams via Chrome session tokens."""

import json
import base64
import time
from typing import Dict, Any, Optional
from common.chrome_cookies import ChromeCookieDecryptor


class TeamsAuthManager:
    _cached_auth: Optional[Dict[str, Any]] = None

    @classmethod
    def get_auth(cls, force_refresh: bool = False) -> Dict[str, Any]:
        now = time.time()
        if not force_refresh and cls._cached_auth:
            if cls._cached_auth.get("exp", 0) > (now + 60):
                return cls._cached_auth

        cookies = ChromeCookieDecryptor.get_cookies_for_domain(
            domain_pattern="asyncgw.teams.microsoft.com",
            cookie_names=["skypetoken_asm"]
        )

        skypetoken = cookies.get("skypetoken_asm")
        if not skypetoken:
            raise RuntimeError("Could not find active 'skypetoken_asm' cookie in Google Chrome.")

        parts = skypetoken.split('.')
        if len(parts) < 2:
            raise RuntimeError("Invalid SkypeToken JWT format.")

        payload_b64 = parts[1]
        payload_b64 += '=' * ((4 - len(payload_b64) % 4) % 4)
        claims = json.loads(base64.urlsafe_b64decode(payload_b64).decode('utf-8', errors='ignore'))

        region = claims.get('rgn', 'apac')
        exp = claims.get('exp', now + 3600)
        skypeid = claims.get('skypeid', '')

        auth_data = {
            "token": skypetoken,
            "region": region,
            "skypeid": skypeid,
            "exp": exp,
            "claims": claims,
            "base_url": f"https://{region}.ng.msg.teams.microsoft.com/v1"
        }
        cls._cached_auth = auth_data
        return auth_data
