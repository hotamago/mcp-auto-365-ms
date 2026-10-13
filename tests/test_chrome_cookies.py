"""Cookie decryption, without touching the real keyring or Chrome profile."""

from __future__ import annotations

import hashlib

from cryptography.hazmat.backends import default_backend
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes

from common.chrome_cookies import ChromeCookieDecryptor


def _encrypt(plaintext: bytes, key: bytes, prefix: bytes, host_hash: bytes = b"") -> bytes:
    data = host_hash + plaintext
    pad = 16 - (len(data) % 16)
    data += bytes([pad]) * pad
    encryptor = Cipher(algorithms.AES(key), modes.CBC(b" " * 16), backend=default_backend()).encryptor()
    return prefix + encryptor.update(data) + encryptor.finalize()


def _key(password: bytes) -> bytes:
    return hashlib.pbkdf2_hmac("sha1", password, b"saltysalt", 1, 16)


def test_v10_uses_the_fixed_peanuts_password(monkeypatch):
    """Chromium encrypts v10 with 'peanuts', not with the keyring secret."""
    monkeypatch.setattr(
        ChromeCookieDecryptor, "get_master_key", classmethod(lambda cls: (_ for _ in ()).throw(AssertionError("keyring used")))
    )
    host = "vingroupjsc.sharepoint.com"
    blob = _encrypt(b"cookie-value", _key(b"peanuts"), b"v10", hashlib.sha256(host.encode()).digest())
    assert ChromeCookieDecryptor.decrypt_value(blob, host) == "cookie-value"


def test_v11_uses_the_keyring_key(monkeypatch):
    secret = _key(b"super-secret")
    monkeypatch.setattr(ChromeCookieDecryptor, "get_master_key", classmethod(lambda cls: secret))
    host = "teams.microsoft.com"
    blob = _encrypt(b"skypetoken-xyz", secret, b"v11", hashlib.sha256(host.encode()).digest())
    assert ChromeCookieDecryptor.decrypt_value(blob, host) == "skypetoken-xyz"


def test_domain_hash_prefix_is_verified_not_blindly_stripped(monkeypatch):
    """A value with no domain-hash prefix must not lose its first 32 bytes."""
    secret = _key(b"k")
    monkeypatch.setattr(ChromeCookieDecryptor, "get_master_key", classmethod(lambda cls: secret))
    plaintext = b"A" * 48  # printable, longer than 32 bytes, no hash prefix
    blob = _encrypt(plaintext, secret, b"v11")
    assert ChromeCookieDecryptor.decrypt_value(blob, "example.com") == "A" * 48


def test_unencrypted_value_passes_through():
    assert ChromeCookieDecryptor.decrypt_value(b"plain-value") == "plain-value"


def test_empty_value():
    assert ChromeCookieDecryptor.decrypt_value(b"") == ""


def _firefox_profile(tmp_path, monkeypatch, expiry: int):
    import sqlite3

    profile = tmp_path / "ab12cd34.default-release"
    profile.mkdir()
    conn = sqlite3.connect(profile / "cookies.sqlite")
    conn.execute("CREATE TABLE moz_cookies (name TEXT, value TEXT, host TEXT, expiry INTEGER, lastAccessed INTEGER)")
    conn.executemany(
        "INSERT INTO moz_cookies VALUES (?, ?, ?, ?, ?)",
        [
            ("FedAuth", "old", ".sharepoint.com", expiry, 1),
            ("FedAuth", "fresh", "vingroupjsc.sharepoint.com", expiry, 2),
            ("other", "x", "example.com", expiry, 3),
        ],
    )
    conn.commit()
    conn.close()
    monkeypatch.setenv("MCP365_BROWSER", "firefox")
    monkeypatch.setenv("MCP365_BROWSER_USER_DATA_DIR", str(tmp_path))
    from common.config import reset_config_cache

    reset_config_cache()
    ChromeCookieDecryptor.clear_cache()


def test_firefox_reads_plain_cookies_without_the_keyring(tmp_path, monkeypatch):
    """Firefox: profile found although browser.profile is still Chromium's 'Default'; newest cookie wins."""
    monkeypatch.setattr(
        ChromeCookieDecryptor, "get_master_key", classmethod(lambda cls: (_ for _ in ()).throw(AssertionError("keyring used")))
    )
    _firefox_profile(tmp_path, monkeypatch, expiry=2_000_000_000)
    assert ChromeCookieDecryptor.get_cookies_for_domain("sharepoint.com", ["FedAuth"]) == {"FedAuth": "fresh"}
    assert ChromeCookieDecryptor.get_cookies_for_domain("example.com", use_cache=False) == {"other": "x"}
    assert ChromeCookieDecryptor.get_cookie_expiry("sharepoint.com", "FedAuth") == 2_000_000_000


def test_firefox_expiry_in_milliseconds(tmp_path, monkeypatch):
    _firefox_profile(tmp_path, monkeypatch, expiry=2_000_000_000_000)
    assert ChromeCookieDecryptor.get_cookie_expiry("sharepoint.com", "FedAuth") == 2_000_000_000
