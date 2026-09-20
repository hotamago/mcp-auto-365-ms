"""Unified MCP server for Microsoft 365: SharePoint, OneDrive and Microsoft Teams.

Run from source via ``uv run``; ``src/`` is put on sys.path first so the flat
intra-project imports resolve.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from mcp.server.mcpserver import MCPServer  # noqa: E402

from tools import register_all  # noqa: E402

mcp = MCPServer("auto-365-ms")
register_all(mcp)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
