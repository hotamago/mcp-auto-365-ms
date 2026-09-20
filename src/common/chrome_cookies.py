"""Chromium-family cookie reader for Linux (libsecret / GNOME Keyring).

Improvements over the original implementation:

* Keyring items are located by their **attributes** (``application=chrome``),
  not by a hardcoded object path like ``/login/2``. That path was an item
  *index*: installing any other Chromium app (Cursor, Termius, …) shifts it and
  silently yields the wrong master key.
* ``v10`` and ``v11`` values use different keys. On Linux Chromium encrypts
  ``v10`` with the fixed password ``peanuts`` and only ``v11`` with the keyring
  secret. Using the keyring key for both is wrong whenever the keyring is
  unavailable and Chrome falls back to basic storage.
* The 32-byte SHA-256 domain prefix that newer Chrome builds prepend is
  detected by verification rather than blindly stripped.
* The cookie DB is copied together with its ``-wal``/``-shm`` sidecars, so
  recently written cookies are not missed, into a private 0700 temp dir that is
  always removed.
"""

from __future__ import annotations

import contextlib
import hashlib
import os
import shutil
import sqlite3
import tempfile
import time
from dataclasses import dataclass
from pathlib import Path

from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

from .config import get_config
from .errors import CookieError, KeyringError

#: Chromium's fallback password when no keyring backend is available.
_BASIC_PASSWORD = b"peanuts"
_SALT = b"saltysalt"
_IV = b" " * 16
_KEY_LENGTH = 16
_ITERATIONS = 1

#: browser name -> (user-data dir candidates, keyring ``application`` attribute)
_BROWSERS: dict[str, tuple[tuple[str, ...], str]] = {
    "chrome": (("~/.config/google-chrome",), "chrome"),
    "chromium": (("~/.config/chromium", "~/snap/chromium/common/chromium"), "chromium"),
    "brave": (("~/.config/BraveSoftware/Brave-Browser",), "brave"),
    "edge": (("~/.config/microsoft-edge",), "microsoft-edge"),
}


def _derive(password: bytes) -> bytes:
    return hashlib.pbkdf2_hmac("sha1", password, _SALT, _ITERATIONS, _KEY_LENGTH)


@dataclass
class _CacheEntry:
    value: dict[str, str]
    created: float


class ChromeCookieDecryptor:
    _keyring_key: bytes | None = None
    _basic_key: bytes | None = None
    _cache: dict[str, _CacheEntry] = {}
    _cache_ttl: float = 15.0

    # ------------------------------------------------------------------ keys

    @classmethod
    def get_master_key(cls) -> bytes:
        """Master key for ``v11`` values, read from libsecret by attributes."""
        if cls._keyring_key is not None:
            return cls._keyring_key

        browser = get_config().browser.name.lower()
        app_attr = _BROWSERS.get(browser, _BROWSERS["chrome"])[1]

        try:
            import secretstorage
        except ImportError as exc:  # pragma: no cover - dependency is declared
            raise KeyringError(
                "Thiếu thư viện 'secretstorage'.",
                "Chạy: uv sync",
            ) from exc

        try:
            bus = secretstorage.dbus_init()
        except Exception as exc:
            raise KeyringError(
                f"Không kết nối được tới D-Bus secret service: {exc}",
                "Đảm bảo bạn đang trong một phiên desktop có gnome-keyring/KWallet đang chạy.",
            ) from exc

        password: bytes | None = None
        locked_collections: list[str] = []
        try:
            for collection in secretstorage.get_all_collections(bus):
                if collection.is_locked():
                    with contextlib.suppress(Exception):
                        locked_collections.append(collection.get_label())
                    continue
                for item in collection.get_all_items():
                    attrs = item.get_attributes() or {}
                    label = (item.get_label() or "").lower()
                    if attrs.get("application", "").lower() == app_attr:
                        password = item.get_secret()
                        break
                    # Fallback for items that omit the application attribute.
                    if not password and label == f"{app_attr} safe storage":
                        password = item.get_secret()
                if password:
                    break
        except Exception as exc:
            raise KeyringError(
                f"Lỗi khi đọc keyring: {exc}",
                "Mở ứng dụng 'Passwords and Keys' (seahorse) và mở khoá keyring 'Login'.",
            ) from exc
        finally:
            with contextlib.suppress(Exception):
                bus.close()

        if not password:
            hint = (
                f"Keyring đang bị KHOÁ: {', '.join(locked_collections)}. Mở khoá rồi thử lại."
                if locked_collections
                else f"Không tìm thấy mục '{app_attr} Safe Storage'. Hãy mở {browser} ít nhất một lần, "
                f"hoặc đặt MCP365_BROWSER cho đúng trình duyệt bạn dùng."
            )
            raise KeyringError(f"Không lấy được master key của {browser} từ keyring.", hint)

        cls._keyring_key = _derive(password)
        return cls._keyring_key

    @classmethod
    def _get_basic_key(cls) -> bytes:
        if cls._basic_key is None:
            cls._basic_key = _derive(_BASIC_PASSWORD)
        return cls._basic_key

    # --------------------------------------------------------------- decrypt

    @classmethod
    def decrypt_value(cls, enc: bytes, host_key: str = "") -> str:
        if not enc:
            return ""
        if isinstance(enc, str):
            enc = enc.encode("utf-8")
        if len(enc) < 4:
            return ""

        prefix, payload = enc[:3], enc[3:]
        if prefix == b"v10":
            key = cls._get_basic_key()
        elif prefix == b"v11":
            key = cls.get_master_key()
        else:
            # Unencrypted (older profiles / no os_crypt).
            return enc.decode("utf-8", errors="ignore")

        if len(payload) % 16:
            payload = payload[: len(payload) - (len(payload) % 16)]

        decryptor = Cipher(algorithms.AES(key), modes.CBC(_IV), backend=default_backend()).decryptor()
        plain = decryptor.update(payload) + decryptor.finalize()

        # Strip PKCS#7 padding.
        if plain:
            pad = plain[-1]
            if 1 <= pad <= 16 and plain[-pad:] == bytes([pad]) * pad:
                plain = plain[:-pad]

        # Newer Chrome builds prepend sha256(host_key). Verify before stripping
        # rather than always cutting 32 bytes.
        if host_key and len(plain) > 32:
            if plain[:32] == hashlib.sha256(host_key.encode("utf-8")).digest():
                plain = plain[32:]
        elif len(plain) > 32:
            # No host to verify against: fall back to the historical behaviour
            # only when the leading bytes are clearly not printable text.
            head = plain[:32]
            if any(b < 0x09 or (0x0E <= b < 0x20) for b in head):
                plain = plain[32:]

        return plain.decode("utf-8", errors="ignore")

    # ------------------------------------------------------------- profiles

    @classmethod
    def _user_data_dir(cls) -> Path:
        cfg = get_config().browser
        if cfg.user_data_dir:
            return Path(os.path.expanduser(cfg.user_data_dir))
        candidates, _ = _BROWSERS.get(cfg.name.lower(), _BROWSERS["chrome"])
        for cand in candidates:
            path = Path(os.path.expanduser(cand))
            if path.is_dir():
                return path
        raise CookieError(
            f"Không tìm thấy thư mục dữ liệu của trình duyệt '{cfg.name}'.",
            "Đặt MCP365_BROWSER_USER_DATA_DIR trỏ tới thư mục profile của trình duyệt.",
        )

    @classmethod
    def _candidate_cookie_dbs(cls) -> list[Path]:
        """Cookie DBs to consult, most-recently-used first."""
        root = cls._user_data_dir()
        configured = get_config().browser.profile

        profiles: list[Path] = []
        if configured and configured.lower() != "auto":
            profiles.append(root / configured)
        else:
            for child in root.iterdir():
                if child.is_dir() and (child.name == "Default" or child.name.startswith("Profile ")):
                    profiles.append(child)
            profiles.sort(key=lambda p: (p / "Cookies").stat().st_mtime if (p / "Cookies").exists() else 0, reverse=True)

        dbs = [p / "Cookies" for p in profiles if (p / "Cookies").exists()]
        if not dbs:
            raise CookieError(
                f"Không tìm thấy file Cookies nào trong {root}.",
                "Kiểm tra MCP365_BROWSER_PROFILE (ví dụ 'Default', 'Profile 1', hoặc 'auto').",
            )
        return dbs

    # ---------------------------------------------------------------- query

    @classmethod
    def get_cookies_for_domain(
        cls,
        domain_pattern: str,
        cookie_names: list[str] | None = None,
        use_cache: bool = True,
    ) -> dict[str, str]:
        """Return decrypted cookies whose host matches ``domain_pattern``."""
        cache_key = f"{domain_pattern}|{','.join(sorted(cookie_names or []))}"
        if use_cache:
            hit = cls._cache.get(cache_key)
            if hit and (time.time() - hit.created) < cls._cache_ttl:
                return dict(hit.value)

        result: dict[str, str] = {}
        for db_path in cls._candidate_cookie_dbs():
            result = cls._read_db(db_path, domain_pattern, cookie_names)
            if result:
                break

        if use_cache:
            cls._cache[cache_key] = _CacheEntry(value=dict(result), created=time.time())
        return result

    @classmethod
    def _read_db(cls, db_path: Path, domain_pattern: str, cookie_names: list[str] | None) -> dict[str, str]:
        tmp_dir = tempfile.mkdtemp(prefix="mcp365-cookies-")
        os.chmod(tmp_dir, 0o700)
        try:
            tmp_db = Path(tmp_dir) / "Cookies"
            shutil.copy2(db_path, tmp_db)
            # Bring along the write-ahead log so freshly written cookies are seen.
            for suffix in ("-wal", "-shm"):
                sidecar = Path(str(db_path) + suffix)
                if sidecar.exists():
                    shutil.copy2(sidecar, Path(str(tmp_db) + suffix))

            query = "SELECT name, host_key, encrypted_value, value, expires_utc FROM cookies WHERE host_key LIKE ?"
            params: list[str] = [f"%{domain_pattern}%"]
            if cookie_names:
                query += f" AND name IN ({','.join('?' * len(cookie_names))})"
                params.extend(cookie_names)
            query += " ORDER BY last_access_utc DESC"

            conn = sqlite3.connect(f"file:{tmp_db}?mode=ro", uri=True)
            try:
                rows = conn.execute(query, params).fetchall()
            finally:
                conn.close()

            out: dict[str, str] = {}
            for name, host_key, enc_val, plain_val, _expires in rows:
                if name in out:
                    continue
                value = cls.decrypt_value(enc_val, host_key) if enc_val else (plain_val or "")
                if value:
                    out[name] = value
            return out
        except sqlite3.DatabaseError as exc:
            raise CookieError(
                f"Không đọc được cookie database ({db_path}): {exc}",
                "Đóng hẳn Chrome rồi thử lại, hoặc kiểm tra quyền truy cập file.",
            ) from exc
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    @classmethod
    def get_cookie_expiry(cls, domain_pattern: str, cookie_name: str) -> float | None:
        """Unix timestamp at which ``cookie_name`` expires, if known."""
        for db_path in cls._candidate_cookie_dbs():
            tmp_dir = tempfile.mkdtemp(prefix="mcp365-cookies-")
            os.chmod(tmp_dir, 0o700)
            try:
                tmp_db = Path(tmp_dir) / "Cookies"
                shutil.copy2(db_path, tmp_db)
                conn = sqlite3.connect(f"file:{tmp_db}?mode=ro", uri=True)
                try:
                    row = conn.execute(
                        "SELECT expires_utc FROM cookies WHERE host_key LIKE ? AND name = ? "
                        "ORDER BY last_access_utc DESC LIMIT 1",
                        (f"%{domain_pattern}%", cookie_name),
                    ).fetchone()
                finally:
                    conn.close()
                if row and row[0]:
                    # Chrome stores microseconds since 1601-01-01.
                    return (row[0] / 1_000_000) - 11_644_473_600
            except Exception:
                continue
            finally:
                shutil.rmtree(tmp_dir, ignore_errors=True)
        return None

    @classmethod
    def clear_cache(cls) -> None:
        cls._cache.clear()
