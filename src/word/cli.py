"""Command-line utility for the Word Companion Add-in and dev bridge."""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from common.config import get_config
from word.bridge import get_bridge
from word.certs import get_or_create_dev_certificate


def cmd_manifest(args: argparse.Namespace) -> None:
    output_path = Path(args.output)
    addin_dir = Path(__file__).resolve().parent / "addin"
    manifest_src = addin_dir / "manifest.xml"
    if not manifest_src.is_file():
        print(f"Error: template manifest not found at {manifest_src}", file=sys.stderr)
        sys.exit(1)

    cfg = get_config().word
    proto = "https" if cfg.ssl_enabled else "http"
    base_url = f"{proto}://{cfg.host}:{cfg.port}"

    text = manifest_src.read_text(encoding="utf-8")
    text = text.replace("https://127.0.0.1:3650", base_url)
    text = text.replace("https://localhost:3650", base_url)

    output_path.write_text(text, encoding="utf-8")
    print(f"✓ Manifest đã xuất tại: {output_path.resolve()}")
    print("  Để sideload vào Word Online: Home > Add-ins > More Settings > Upload My Add-in.")


def cmd_cert(args: argparse.Namespace) -> None:
    cert_path, key_path = get_or_create_dev_certificate()
    print(f"✓ Chứng chỉ dev SSL:\n  - Cert: {cert_path.resolve()}\n  - Key:  {key_path.resolve()}")


def cmd_status(args: argparse.Namespace) -> None:
    cfg = get_config().word
    bridge = get_bridge()
    bridge.ensure_running(
        host=cfg.host,
        port=cfg.port,
        ssl_enabled=cfg.ssl_enabled,
        cert_file=cfg.cert_file,
        key_file=cfg.key_file,
    )
    status = bridge.get_status_summary()
    sessions = bridge.get_active_sessions_info()
    print(f"Word Companion Bridge: {status['status']} ({status['base_url']})")
    print(f"Phiên tài liệu đang kết nối: {len(sessions)}")
    for s in sessions:
        print(f"  - {s['doc_title']} ({s['session_id']}): {s['doc_url']}")


def cmd_start(args: argparse.Namespace) -> None:
    cfg = get_config().word
    bridge = get_bridge()
    bridge.ensure_running(
        host=cfg.host,
        port=cfg.port,
        ssl_enabled=cfg.ssl_enabled,
        cert_file=cfg.cert_file,
        key_file=cfg.key_file,
    )
    print(f"Word Companion Bridge đang chạy tại {bridge.base_url}")
    print("Nhấn Ctrl+C để dừng.")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        bridge.stop()
        print("\nĐã dừng.")


def main() -> None:
    parser = argparse.ArgumentParser(description="MCP Auto 365 MS - Word Companion CLI")
    subparsers = parser.add_subparsers(dest="subcommand", required=True)

    p_manifest = subparsers.add_parser("manifest", help="Xuất file manifest.xml để sideload vào Word")
    p_manifest.add_argument("-o", "--output", default="mcp-auto-365-word-manifest.xml", help="Đường dẫn file đích")
    p_manifest.set_defaults(func=cmd_manifest)

    p_cert = subparsers.add_parser("cert", help="Tạo hoặc kiểm tra chứng chỉ dev TLS localhost")
    p_cert.set_defaults(func=cmd_cert)

    p_status = subparsers.add_parser("status", help="Kiểm tra trạng thái bridge và tài liệu đang kết nối")
    p_status.set_defaults(func=cmd_status)

    p_start = subparsers.add_parser("start", help="Khởi động bridge ở chế độ standalone")
    p_start.set_defaults(func=cmd_start)

    parsed = parser.parse_args()
    parsed.func(parsed)


if __name__ == "__main__":
    main()
