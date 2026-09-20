"""Shared Chrome cookie decryption utility for Linux using GNOME Keyring."""

import dbus
import hashlib
import sqlite3
import shutil
import os
from typing import Dict, Optional, List
from cryptography.hazmat.primitives.ciphers import Cipher, algorithms, modes
from cryptography.hazmat.backends import default_backend


class ChromeCookieDecryptor:
    _master_key: Optional[bytes] = None

    @classmethod
    def get_master_key(cls) -> bytes:
        if cls._master_key:
            return cls._master_key

        bus = dbus.SessionBus()
        service = bus.get_object('org.freedesktop.secrets', '/org/freedesktop/secrets')
        iface = dbus.Interface(service, 'org.freedesktop.Secret.Service')
        output, session_path = iface.OpenSession('plain', dbus.String('', variant_level=1))
        secrets = iface.GetSecrets([dbus.ObjectPath('/org/freedesktop/secrets/collection/login/2')], session_path)
        password = bytes(list(secrets.values())[0][2])
        key = hashlib.pbkdf2_hmac('sha1', password, b'saltysalt', 1, 16)
        cls._master_key = key
        return key

    @classmethod
    def decrypt_value(cls, enc: bytes) -> str:
        if not enc:
            return ''
        if isinstance(enc, str):
            enc = enc.encode('utf-8')
        if len(enc) < 3:
            return ''

        if enc.startswith(b'v10') or enc.startswith(b'v11'):
            key = cls.get_master_key()
            iv = b' ' * 16
            raw = enc[3:]
            cipher = Cipher(algorithms.AES(key), modes.CBC(iv), backend=default_backend())
            dec = cipher.decryptor().update(raw) + cipher.decryptor().finalize()
            pad = dec[-1]
            if 1 <= pad <= 16:
                dec = dec[:-pad]
            return dec[32:].decode('utf-8', errors='ignore')
        return enc.decode('utf-8', errors='ignore')

    @classmethod
    def get_cookies_for_domain(cls, domain_pattern: str, cookie_names: Optional[List[str]] = None) -> Dict[str, str]:
        """Fetch and decrypt cookies for a domain pattern."""
        cookie_src = os.path.expanduser('~/.config/google-chrome/Default/Cookies')
        if not os.path.exists(cookie_src):
            raise FileNotFoundError(f"Chrome cookies database not found at {cookie_src}")

        tmp_db = f"/tmp/cookies_decrypt_{os.getpid()}.sqlite"
        shutil.copy2(cookie_src, tmp_db)
        conn = sqlite3.connect(tmp_db)
        c = conn.cursor()

        query = "SELECT name, host_key, encrypted_value FROM cookies WHERE host_key LIKE ?"
        params = [f"%{domain_pattern}%"]
        if cookie_names:
            placeholders = ",".join(["?"] * len(cookie_names))
            query += f" AND name IN ({placeholders})"
            params.extend(cookie_names)

        query += " ORDER BY last_access_utc DESC"
        c.execute(query, params)
        rows = c.fetchall()
        conn.close()
        try:
            os.remove(tmp_db)
        except Exception:
            pass

        result = {}
        for name, host_key, enc_val in rows:
            if name not in result:
                val = cls.decrypt_value(enc_val)
                if val:
                    result[name] = val

        return result
