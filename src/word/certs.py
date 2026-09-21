"""Local TLS certificate generator for the Word Companion Add-in dev bridge.

Word on the web (and modern Office web clients) requires HTTPS when loading an
iframe task pane from localhost. This module generates a local self-signed dev
certificate with Subject Alternative Names for ``localhost`` and ``127.0.0.1``,
using the existing ``cryptography`` dependency.
"""

from __future__ import annotations

import datetime
import ipaddress
import os
from pathlib import Path

from cryptography import x509
from cryptography.hazmat.primitives import hashes, serialization
from cryptography.hazmat.primitives.asymmetric import rsa
from cryptography.x509.oid import ExtendedKeyUsageOID, NameOID

_DEFAULT_CERT_DIR = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "mcp-auto-365-ms" / "certs"


def get_or_create_dev_certificate(cert_dir: Path | None = None) -> tuple[Path, Path]:
    """Return paths to ``(cert_path, key_path)``, creating them if needed.

    If ``~/.office-addin-dev-certs/localhost.crt`` exists (from standard Office
    dev tools), it is reused automatically. Otherwise, a 365-day self-signed
    certificate is created in ``cert_dir`` (defaulting to
    ``~/.config/mcp-auto-365-ms/certs/``).
    """
    office_dev_dir = Path.home() / ".office-addin-dev-certs"
    if (office_dev_dir / "localhost.crt").is_file() and (office_dev_dir / "localhost.key").is_file():
        return office_dev_dir / "localhost.crt", office_dev_dir / "localhost.key"

    target_dir = cert_dir or _DEFAULT_CERT_DIR
    target_dir.mkdir(parents=True, exist_ok=True)
    cert_path = target_dir / "localhost.crt"
    key_path = target_dir / "localhost.key"

    if cert_path.is_file() and key_path.is_file():
        return cert_path, key_path

    key = rsa.generate_private_key(public_exponent=65537, key_size=2048)
    subject = x509.Name([
        x509.NameAttribute(NameOID.COMMON_NAME, "localhost"),
        x509.NameAttribute(NameOID.ORGANIZATION_NAME, "MCP Auto 365 MS Dev"),
    ])
    now = datetime.datetime.now(datetime.UTC)
    cert = (
        x509.CertificateBuilder()
        .subject_name(subject)
        .issuer_name(subject)
        .public_key(key.public_key())
        .serial_number(x509.random_serial_number())
        .not_valid_before(now - datetime.timedelta(days=1))
        .not_valid_after(now + datetime.timedelta(days=365))
        .add_extension(
            x509.SubjectAlternativeName([
                x509.DNSName("localhost"),
                x509.IPAddress(ipaddress.IPv4Address("127.0.0.1")),
            ]),
            critical=False,
        )
        .add_extension(
            x509.BasicConstraints(ca=False, path_length=None),
            critical=True,
        )
        .add_extension(
            x509.KeyUsage(
                digital_signature=True,
                content_commitment=False,
                key_encipherment=True,
                data_encipherment=False,
                key_agreement=False,
                key_cert_sign=False,
                crl_sign=False,
                encipher_only=False,
                decipher_only=False,
            ),
            critical=True,
        )
        .add_extension(
            x509.ExtendedKeyUsage([ExtendedKeyUsageOID.SERVER_AUTH]),
            critical=False,
        )
        .sign(key, hashes.SHA256())
    )

    key_bytes = key.private_bytes(
        encoding=serialization.Encoding.PEM,
        format=serialization.PrivateFormat.TraditionalOpenSSL,
        encryption_algorithm=serialization.NoEncryption(),
    )
    cert_bytes = cert.public_bytes(serialization.Encoding.PEM)

    # Write key with 0600 permissions
    key_path.touch(mode=0o600, exist_ok=True)
    key_path.write_bytes(key_bytes)
    cert_path.write_bytes(cert_bytes)

    return cert_path, key_path
