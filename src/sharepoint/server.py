"""Standalone MCP server exposing only the SharePoint/OneDrive tools."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from mcp.server.mcpserver import MCPServer  # noqa: E402

from tools import register_resources, register_sharepoint_tools  # noqa: E402

mcp = MCPServer("doc-reader")
register_sharepoint_tools(mcp)
register_resources(mcp)


def main() -> None:
    from common.config import get_config

    cfg = get_config().word
    if cfg.enabled:
        try:
            from word.bridge import get_bridge

            get_bridge().ensure_running(
                host=cfg.host,
                port=cfg.port,
                ssl_enabled=cfg.ssl_enabled,
                cert_file=cfg.cert_file,
                key_file=cfg.key_file,
            )
        except Exception as exc:
            import logging

            logging.getLogger("doc-reader").warning("Could not start Word Companion Bridge: %s", exc)
    mcp.run()

if __name__ == "__main__":
    main()
