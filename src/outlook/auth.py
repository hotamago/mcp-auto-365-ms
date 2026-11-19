"""Outlook browser-session authentication without Graph consent.

The configured Chromium profile already carries persistent Microsoft sign-in
cookies. We replay those cookies only to ``login.microsoftonline.com`` and use
Outlook Web's own first-party SPA authorization (the same flow the browser
runs) to mint an ``https://outlook.office.com`` token. No device login, app
registration, Graph grant, refresh-token cache, or tenant-admin permission is
introduced. The short-lived access token stays in memory. The flow itself is
:class:`common.spa_auth.SpaTokenManager`, shared with the Teams sections channel.
"""

from __future__ import annotations

from common.config import get_config
from common.errors import AuthExpiredError, Mcp365Error
from common.spa_auth import SpaSettings, SpaTokenManager
from teams.auth import TeamsAuthManager


class MailAuthManager(SpaTokenManager):
    """Mint and cache Outlook Web's short-lived token from browser cookies."""

    def _settings(self) -> SpaSettings:
        cfg = get_config().mail
        origin = cfg.origin.rstrip("/")
        return SpaSettings(
            product="Outlook",
            client_id=cfg.client_id,
            tenant_id=cfg.tenant_id,
            login_host=cfg.login_host,
            origin=origin,
            scope=cfg.scope,
            redirect_uri=cfg.redirect_uri,
            audience=origin,
            open_url=origin,
        )

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
