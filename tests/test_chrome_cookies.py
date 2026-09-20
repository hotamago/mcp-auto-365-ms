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
