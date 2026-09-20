"""Standalone MCP server exposing only the Microsoft Teams tools."""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from mcp.server.mcpserver import MCPServer  # noqa: E402

from tools import register_prompts, register_resources, register_teams_tools  # noqa: E402

mcp = MCPServer("teams-reader")
register_teams_tools(mcp)
register_resources(mcp)
register_prompts(mcp)


def main() -> None:
    mcp.run()


if __name__ == "__main__":
    main()
