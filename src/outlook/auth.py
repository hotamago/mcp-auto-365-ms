"""Delegated Microsoft Graph authentication for Outlook mail.

Mail uses a separate MSAL token because the Azure CLI first-party client is not
pre-authorized for ``Mail.Read`` or ``Mail.Send``. Device-code login grants only
those delegated scopes for the signed-in user's mailbox; no application
permission or tenant-admin Graph grant is required unless tenant policy disables
user consent.
"""

from __future__ import annotations

import json
import os
import threading
from pathlib import Path
from typing import Any

import msal

from common.config import get_config
from common.errors import AuthExpiredError, Mcp365Error

_MAIL_SCOPES = ["Mail.Read", "Mail.Send"]
_GRAPH_CLI_CLIENT_ID = "14d82eec-204b-4c2f-b7e8-296a70dab67e"


class MailAuthManager:
    """Own the MSAL device flow and a private, persistent token cache."""

    def __init__(self, cache_path: Path | None = None) -> None:
        self._cache_path_override = cache_path
        self._cache = msal.SerializableTokenCache()
        self._app: msal.PublicClientApplication | None = None
        self._loaded = False
        self._lock = threading.RLock()
        self._flow: dict[str, Any] | None = None
        self._thread: threading.Thread | None = None
        self._result: dict[str, Any] | None = None

    @property
    def cache_path(self) -> Path:
        if self._cache_path_override is not None:
            return self._cache_path_override
        config_home = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config"))
        return config_home / "mcp-auto-365-ms" / "mail-token-cache.json"

    def _load_cache(self) -> None:
        if self._loaded:
            return
        self._loaded = True
        try:
            serialized = self.cache_path.read_text(encoding="utf-8")
        except FileNotFoundError:
            return
        except OSError as exc:
            raise Mcp365Error(
                f"Không đọc được cache đăng nhập mail `{self.cache_path}`: {exc}",
                "Kiểm tra quyền thư mục ~/.config/mcp-auto-365-ms rồi thử lại.",
            ) from exc
        if serialized:
            try:
                self._cache.deserialize(serialized)
            except (ValueError, json.JSONDecodeError):
                # A truncated cache cannot authenticate and contains no durable
                # user data. Ignore it; the next login atomically replaces it.
                pass

    def _save_cache(self) -> None:
        if not self._cache.has_state_changed:
            return
        path = self.cache_path
        path.parent.mkdir(mode=0o700, parents=True, exist_ok=True)
        os.chmod(path.parent, 0o700)
        tmp = path.with_name(f".{path.name}.tmp")
        tmp.write_text(self._cache.serialize(), encoding="utf-8")
        os.chmod(tmp, 0o600)
        tmp.replace(path)

    def _application(self) -> msal.PublicClientApplication:
        with self._lock:
            self._load_cache()
            if self._app is None:
                cfg = get_config().mail
                client_id = cfg.client_id.strip() or _GRAPH_CLI_CLIENT_ID
                tenant = cfg.tenant_id.strip() or "organizations"
                self._app = msal.PublicClientApplication(
                    client_id=client_id,
                    authority=f"https://login.microsoftonline.com/{tenant}",
                    token_cache=self._cache,
                )
            return self._app

    @staticmethod
    def _error(result: dict[str, Any]) -> AuthExpiredError:
        code = result.get("error", "authentication_failed")
        detail = str(result.get("error_description", ""))[:500]
        return AuthExpiredError(
            f"Không đăng nhập được Outlook mail ({code})." + (f"\n{detail}" if detail else ""),
            "Gọi `start_mail_login`, mở URL được trả về, nhập mã và đăng nhập. Nếu tenant chặn user consent, "
            "nhờ admin cho phép delegated Mail.Read và Mail.Send cho ứng dụng.",
        )

    def _silent_result(self) -> dict[str, Any] | None:
        app = self._application()
        accounts = app.get_accounts()
        if not accounts:
            return None
        result = app.acquire_token_silent(_MAIL_SCOPES, account=accounts[0])
        self._save_cache()
        return result

    def get_token(self) -> str:
        """Return a delegated mail token, never starting interactive auth implicitly."""
        with self._lock:
            result = self._silent_result()
            if result and result.get("access_token"):
                return str(result["access_token"])
            pending = self._thread is not None and self._thread.is_alive()
            if result and result.get("error"):
                raise self._error(result)
        hint = "Đăng nhập đang chờ hoàn tất; gọi `check_mail_login` sau khi nhập mã." if pending else (
            "Gọi `start_mail_login`, làm theo URL + mã trả về, rồi gọi `check_mail_login`."
        )
        raise AuthExpiredError("Outlook mail chưa có phiên đăng nhập delegated hợp lệ.", hint)

    def start_device_login(self) -> dict[str, Any]:
        """Start non-blocking device-code auth and return the exact instructions."""
        with self._lock:
            existing = self._silent_result()
            if existing and existing.get("access_token"):
                claims = existing.get("id_token_claims") or {}
                return {
                    "status": "connected",
                    "username": claims.get("preferred_username") or claims.get("name") or "",
                }

            if self._thread is not None and self._thread.is_alive() and self._flow:
                return self._flow_preview("pending")

            app = self._application()
            flow = app.initiate_device_flow(scopes=_MAIL_SCOPES)
            if "user_code" not in flow:
                raise self._error(flow)
            self._flow = flow
            self._result = None
            self._thread = threading.Thread(target=self._complete_device_flow, args=(app, flow), daemon=True)
            self._thread.start()
            return self._flow_preview("pending")

    def _flow_preview(self, status: str) -> dict[str, Any]:
        flow = self._flow or {}
        return {
            "status": status,
            "verification_uri": flow.get("verification_uri") or flow.get("verification_uri_complete") or "",
            "user_code": flow.get("user_code", ""),
            "expires_in": flow.get("expires_in", 0),
            "message": flow.get("message", ""),
        }

    def _complete_device_flow(self, app: msal.PublicClientApplication, flow: dict[str, Any]) -> None:
        result = app.acquire_token_by_device_flow(flow)
        with self._lock:
            self._result = result
            self._save_cache()

    def login_status(self) -> dict[str, Any]:
        """Report device-flow state without exposing tokens or starting a flow."""
        with self._lock:
            if self._thread is not None and self._thread.is_alive():
                return self._flow_preview("pending")
            if self._result is not None:
                result = self._result
                if result.get("access_token"):
                    claims = result.get("id_token_claims") or {}
                    return {
                        "status": "connected",
                        "username": claims.get("preferred_username") or claims.get("name") or "",
                    }
                return {
                    "status": "failed",
                    "error": result.get("error", "authentication_failed"),
                    "detail": str(result.get("error_description", ""))[:500],
                }
            silent = self._silent_result()
            if silent and silent.get("access_token"):
                claims = silent.get("id_token_claims") or {}
                return {
                    "status": "connected",
                    "username": claims.get("preferred_username") or claims.get("name") or "",
                }
            return {"status": "not_connected"}
