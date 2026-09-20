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
    mcp.run()


if __name__ == "__main__":
    main()
