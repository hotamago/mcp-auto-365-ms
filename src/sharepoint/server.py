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

@mcp.tool()
def upload_sharepoint_file(local_file_path: str, target_folder_url_or_path: str, target_file_name: str = "") -> str:
    """Upload a local file to a SharePoint folder (creates new file or replaces in-place).
    
    Args:
        local_file_path: Path to the local file to upload (e.g. 'docs/report.xlsx', 'output.pdf').
        target_folder_url_or_path: Full SharePoint folder URL or relative path inside site (e.g. 'VF-VSF Collaboration/00.ViTa/...').
        target_file_name: Optional custom remote filename (defaults to local file name).
    """
    try:
        res = sp_client.upload_file(local_file_path, target_folder_url_or_path, target_file_name if target_file_name else None)
        sz = res['size']
        sz_str = f"{sz / (1024*1024):.2f} MB" if sz > 1024*1024 else f"{sz / 1024:.1f} KB"
        return (
            f"✓ Successfully uploaded `{res['name']}` ({sz_str}) to SharePoint!\n"
            f"- **Target Folder:** `{res['folder']}`\n"
            f"- **Item ID:** `{res['id']}`\n"
            f"- **Web URL:** {res['webUrl']}"
        )
    except Exception as e:
        return f"Error uploading file to SharePoint: {e}"


@mcp.tool()
def replace_sharepoint_file(local_file_path: str, file_url_or_guid: str) -> str:
    """Replace an existing SharePoint file with a new version from local disk.
    
    Creates a new version in SharePoint's version history while keeping the same file link and ID.
    
    Args:
        local_file_path: Path to the local updated file.
        file_url_or_guid: SharePoint file URL, sharing link, or document unique GUID.
    """
    try:
        res = sp_client.replace_file(local_file_path, file_url_or_guid)
        sz = res['size']
        sz_str = f"{sz / (1024*1024):.2f} MB" if sz > 1024*1024 else f"{sz / 1024:.1f} KB"
        return (
            f"✓ Successfully replaced `{res['name']}` ({sz_str}) on SharePoint!\n"
            f"- **New Version:** `{res['version']}`\n"
            f"- **Item ID:** `{res['id']}`\n"
            f"- **Modified Time:** {res['modified']}\n"
            f"- **Web URL:** {res['webUrl']}"
        )
    except Exception as e:
        return f"Error replacing SharePoint file: {e}"


@mcp.tool()
def search_sharepoint_files(query: str, max_results: int = 20, file_extension: str = "") -> str:
    """Search across SharePoint files and folders by keyword or file type.
    
    Args:
        query: Keyword to search for (e.g. 'SYS2', 'CAN', 'Architecture', 'DTC').
        max_results: Maximum number of files to return (default: 20).
        file_extension: Optional file extension filter (e.g. 'docx', 'xlsx', 'pptx', 'pdf').
    """
    try:
        results = sp_client.search_files(query=query, max_results=max_results, file_extension=file_extension if file_extension else None)
        if not results:
            return f"No SharePoint documents found matching query '{query}'."

        out = [f"# SharePoint Search Results for '{query}' ({len(results)} found)\n"]
        out.append("| Document Name | Size | Last Modified | Author | UniqueId / Download Link |")
        out.append("| --- | --- | --- | --- | --- |")
        for r in results:
            sz = r['size']
            sz_str = f"{sz / (1024*1024):.2f} MB" if sz > 1024*1024 else f"{sz / 1024:.1f} KB"
            uid = r['unique_id']
            out.append(f"| **[{r['title']}]({r['path']})** | {sz_str} | {r['modified'][:10]} | {r['author'][:25]} | `{uid}` |")

        out.append("\n> **Download Hint:** Call `download_sharepoint_link(unique_id)` or `download_sharepoint_link(path)` to download any file.")
        return "\n".join(out)
    except Exception as e:
        return f"Error searching SharePoint files: {e}"

@mcp.tool()
def compare_sharepoint_versions(file_a: str, file_b: str = "", version_a: str = "", version_b: str = "") -> str:
    """Compare two SharePoint document versions or compare a local file against a SharePoint document.
    
    Outputs a clean Markdown diff / changelog with added, removed, and modified lines.
    Supports Word documents (.docx), text files (.txt, .md, .py, .csv, .json, .yaml), and spreadsheet overview (.xlsx).
    
    Args:
        file_a: Local file path or SharePoint URL/GUID.
        file_b: Optional second file path or SharePoint URL/GUID to compare against file_a.
        version_a: (If file_b omitted) Earlier version label (e.g. '1.0').
        version_b: (If file_b omitted) Later version label (e.g. '2.0', 'latest').
    """
    try:
        if file_b:
            return sp_client.compare_documents(file_a, file_b)
        else:
            return sp_client.compare_versions(file_a, version_a=version_a, version_b=version_b)
    except Exception as e:
        return f"Error comparing document versions: {e}"
if __name__ == "__main__":
    mcp.run()
