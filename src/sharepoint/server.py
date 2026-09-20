"""Standalone MCP Server for exploring and downloading SharePoint/OneDrive links."""

import sys
from pathlib import Path

# Ensure src root is in python path
src_root = str(Path(__file__).resolve().parent.parent)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

from mcp.server.mcpserver import MCPServer
from sharepoint.client import SharePointClient

mcp = MCPServer("doc-reader")
sp_client = SharePointClient()


@mcp.tool()
def read_sharepoint_link(url: str, max_depth: int = 2) -> str:
    """Explore a SharePoint folder or view document metadata.
    
    - If folder: returns folder hierarchy, child files, sizes, modified dates, authors.
    - If document: returns metadata (path, size, author, created/modified date, version history).
    """
    try:
        return sp_client.read_link(url, max_depth=max_depth)
    except Exception as e:
        return f"Error exploring SharePoint link: {e}"


@mcp.tool()
def download_sharepoint_link(url_or_guid: str, target_dir: str = "docs/sharepoint") -> str:
    """Download the original raw file or an entire folder from SharePoint into target_dir.
    
    - Preserves exact binary format (.docx, .xlsx, .pptx, .pdf, images, charts).
    - If folder: recursively downloads all documents preserving subfolder structure.
    - If document: downloads that specific file directly.
    """
    try:
        return sp_client.download_link(url_or_guid, target_dir=target_dir)
    except Exception as e:
        return f"Error downloading from SharePoint: {e}"


if __name__ == "__main__":
    mcp.run()
