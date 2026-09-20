"""Unified MCP Server for Microsoft 365: SharePoint, OneDrive & MS Teams."""

import sys
from pathlib import Path

# Ensure src root is in python path
src_root = str(Path(__file__).resolve().parent)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

from mcp.server.mcpserver import MCPServer
from sharepoint.client import SharePointClient
from teams.client import TeamsClient

# Initialize unified MCP server
mcp = MCPServer("auto-365-ms")
sp_client = SharePointClient()
teams_client = TeamsClient()


# ==========================================
# 1. SharePoint & OneDrive Tools
# ==========================================

@mcp.tool()
def read_sharepoint_link(url: str, max_depth: int = 2) -> str:
    """Explore a SharePoint folder or view document metadata.
    
    - If folder: returns folder hierarchy, child files, sizes, modified dates, authors.
    - If document: returns metadata (SharePoint path, size, author, created/modified date, version history). Does NOT convert to markdown.
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

# ==========================================
# 2. Microsoft Teams Automation Tools
# ==========================================

@mcp.tool()
def get_recent_team_messages(hours: int = 48, max_chats: int = 8, limit_per_chat: int = 8, filter_keyword: str = "") -> str:
    """Scan and fetch new messages across all active group chats in a single parallel call.
    
    Eliminates the need to call list_chats and read each chat one by one.
    
    Args:
        hours: How many hours back to look (default: 48).
        max_chats: Maximum active chats to scan (default: 8).
        limit_per_chat: Maximum recent messages per chat (default: 8).
        filter_keyword: Optional filter for chat names or content.
    """
    try:
        feed = teams_client.get_recent_feed(hours=hours, max_chats=max_chats, limit_per_chat=limit_per_chat, filter_keyword=filter_keyword)
        if not feed:
            return f"No active messages found in the last {hours} hours."

        out = [f"# Teams Recent Activity Feed (Past {hours} hours — {len(feed)} active chats)\n"]
        for f in feed:
            out.append(f"## 💬 {f['chat_name']}")
            out.append(f"- **Chat ID:** `{f['chat_id']}`")
            out.append(f"- **Messages:** {len(f['messages'])}\n")
            for m in f['messages']:
                sp_links = m.get("sharepoint_links", [])
                links_text = ""
                if sp_links:
                    links_text = "\n  *Shared Links:* " + ", ".join([f"[SharePoint Link]({l})" for l in sp_links])
                out.append(f"- **[{m['timestamp']}] {m['sender']}:** {m['content']}{links_text}")
            out.append("\n---")
        return "\n".join(out)
    except Exception as e:
        return f"Error fetching recent team feed: {e}"


@mcp.tool()
def get_my_mentions(hours: int = 72, limit: int = 20, context_before: int = 2, context_after: int = 2) -> str:
    """Find all messages specifically mentioning you (@Nguyễn Hoàng Sơn, @Sơn, @all) across all group chats.
    
    Includes a surrounding discussion context window (messages before and after the mention) so you can
    fully understand the context, background discussion, links, and requirements of the assigned tasks.
    
    Args:
        hours: How many hours back to look (default: 72).
        limit: Maximum number of mentions to return (default: 20).
        context_before: Number of preceding messages before the mention to include (default: 2, 0 to disable).
        context_after: Number of following messages after the mention to include (default: 2, 0 to disable).
    """
    try:
        mentions = teams_client.get_user_mentions(
            hours=hours,
            limit=limit,
            context_before=max(0, min(context_before, 10)),
            context_after=max(0, min(context_after, 10))
        )
        if not mentions:
            return f"No mentions found in the last {hours} hours."

        out = [f"# Messages Mentioning You ({len(mentions)} found in past {hours} hours)\n"]
        for m in mentions:
            out.append(f"### 📍 [{m['chat_name']}] — Tagged by **{m['sender']}** ({m['timestamp']})")
            
            ctx = m.get("context", [])
            if ctx and (context_before > 0 or context_after > 0):
                out.append("\n**Discussion Thread Context:**")
                for c in ctx:
                    time_part = c['timestamp'][11:19] if len(c['timestamp']) >= 19 else c['timestamp']
                    if c["is_mention"]:
                        out.append(f"👉 **[{time_part}] {c['sender']} (MENTION):**")
                        out.append(f"> {c['content']}")
                        if c.get("sharepoint_links"):
                            out.append(f"> *Links:* " + ", ".join([f"[Link]({l})" for l in c["sharepoint_links"]]))
                    else:
                        rel_pos = f"{c['offset']:+d}"
                        out.append(f"- *({rel_pos}) [{time_part}] {c['sender']}:* {c['content']}")
                        if c.get("sharepoint_links"):
                            out.append(f"  *Links:* " + ", ".join([f"[Link]({l})" for l in c["sharepoint_links"]]))
            else:
                out.append(f"> {m['content']}")
                if m.get("sharepoint_links"):
                    out.append(f"> *Links:* " + ", ".join(m["sharepoint_links"]))
            
            out.append("\n---\n")
        return "\n".join(out)
    except Exception as e:
        return f"Error retrieving mentions: {e}"

@mcp.tool()
def send_teams_message(chat_name_or_id: str, message: str) -> str:
    """Send a message to a specific Teams group chat or 1:1 chat when commanded.
    
    Args:
        chat_name_or_id: Exact or partial name of the chat (e.g. 'Proactive Agent', 'ViTa S5', 'Back-end') or the chat thread ID.
        message: The message text to send (markdown formatting like **bold** and `code` is supported).
    """
    try:
        res = teams_client.send_message(conversation_id_or_name=chat_name_or_id, message=message)
        return f"✓ Message sent successfully to **{res['conversation_name']}** (`{res['conversation_id']}`):\n\n> {res['message_sent']}"
    except Exception as e:
        return f"Error sending message to Teams: {e}"


@mcp.tool()
def download_chat_attachments(chat_name_or_id: str, target_dir: str = "docs/sharepoint", limit: int = 5) -> str:
    """Automatically discover SharePoint/OneDrive links posted in a Teams chat and download the files.
    
    Args:
        chat_name_or_id: Exact or partial chat name (e.g. 'Proactive Agent', 'ViTa S5') or chat thread ID.
        target_dir: Directory where files should be downloaded (default: 'docs/sharepoint').
        limit: Maximum number of recent files to download (default: 5).
    """
    try:
        res = teams_client.get_messages(chat_name_or_id, limit=30)
        messages = res.get("messages", [])
        links = []
        for m in messages:
            for link in m.get("sharepoint_links", []):
                if link not in links:
                    links.append(link)

        if not links:
            return f"No SharePoint/OneDrive links found in the recent messages of '{res.get('conversation_name')}'."

        downloaded_reports = []
        for l in links[:limit]:
            dl_res = sp_client.download_link(l, target_dir=target_dir)
            downloaded_reports.append(dl_res)

        return f"# Downloaded {len(downloaded_reports)} attachments from chat '{res.get('conversation_name')}':\n\n" + "\n\n---\n\n".join(downloaded_reports)
    except Exception as e:
        return f"Error downloading chat attachments: {e}"


@mcp.tool()
def read_teams_chat(chat_name_or_id: str, limit: int = 30, since: str = "", only_mentions: bool = False) -> str:
    """Read full message history and discussions from a specific Teams group chat or 1:1 chat.
    
    Args:
        chat_name_or_id: The exact or partial name of the chat (e.g. 'Proactive Agent', 'ViTa S5') or the chat thread ID.
        limit: Number of recent messages to fetch (default: 30, max: 100).
        since: Optional filter for messages since a date/time (e.g. 'today', 'yesterday', '2026-09-18').
        only_mentions: If True, only returns messages that mention the user.
    """
    try:
        res = teams_client.get_messages(chat_name_or_id, limit=limit, since=since if since else None, only_mentions=only_mentions)
        messages = res.get("messages", [])
        if not messages:
            return f"No messages found for chat '{res.get('conversation_name')}'."

        out = [
            f"# Chat: {res.get('conversation_name')}",
            f"- **Thread ID:** `{res.get('conversation_id')}`",
            f"- **Messages Retrieved:** {len(messages)}\n",
            "---"
        ]

        for m in messages:
            sp_links = m.get("sharepoint_links", [])
            links_text = ""
            if sp_links:
                links_text = "\n  *SharePoint Links:* " + ", ".join([f"[Link]({l})" for l in sp_links])

            out.append(f"### [{m['timestamp']}] {m['sender']}\n{m['content']}{links_text}\n")

        return "\n".join(out)
    except Exception as e:
        return f"Error reading Teams chat: {e}"


@mcp.tool()
def list_teams_chats(limit: int = 30, filter_keyword: str = "") -> str:
    """List recent Microsoft Teams group chats, direct chats, and meeting chats.
    
    Args:
        limit: Maximum number of chats to return (default: 30).
        filter_keyword: Optional keyword to filter chat names or last messages.
    """
    try:
        chats = teams_client.list_conversations(page_size=limit, filter_keyword=filter_keyword)
        if not chats:
            return "No conversations found matching the criteria."

        out = [f"# Microsoft Teams Conversations ({len(chats)})\n"]
        out.append("| Type | Chat Name | Last Sender | Last Activity | Chat ID |")
        out.append("| --- | --- | --- | --- | --- |")
        for c in chats:
            time_str = c["last_activity"][:19].replace("T", " ") if c["last_activity"] else "N/A"
            out.append(f"| {c['type']} | **{c['name']}** | {c['last_sender']} | {time_str} | `{c['id']}` |")
            if c['last_message']:
                out.append(f"> *Latest:* \"{c['last_message'][:120]}\"")
        return "\n".join(out)
    except Exception as e:
        return f"Error listing Teams chats: {e}"


@mcp.tool()
def search_teams_chat_messages(query: str, limit: int = 20) -> str:
    """Search for keywords, technical discussions, bug reports, or task requests across all recent group chats.
    
    Args:
        query: Keyword or phrase to search for (e.g. 'DTC', 'S5', 'B1024', 'review', 'QC').
        limit: Maximum number of matching messages to return (default: 20).
    """
    try:
        hits = teams_client.search_messages(query=query, max_results=limit)
        if not hits:
            return f"No messages found containing keyword: '{query}'."

        out = [f"# Search Results for '{query}' ({len(hits)} hits)\n"]
        for h in hits:
            out.append(f"**[{h['chat_name']}]** — *{h['sender']}* ({h['timestamp']}):")
            out.append(f"> {h['content'][:300]}")
            if h.get("sharepoint_links"):
                out.append(f"> *Links:* " + ", ".join(h["sharepoint_links"]))
            out.append("")
        return "\n".join(out)
    except Exception as e:
        return f"Error searching Teams messages: {e}"


@mcp.tool()
def edit_teams_message(chat_name_or_id: str, message_id: str, new_message: str) -> str:
    """Edit an existing message sent by you in a Teams group chat or 1:1 chat.
    
    Args:
        chat_name_or_id: Name or thread ID of the chat.
        message_id: ID of the message to edit.
        new_message: The updated message content.
    """
    try:
        res = teams_client.edit_message(chat_name_or_id, message_id=message_id, new_message=new_message)
        return f"✓ Successfully edited message `{res['message_id']}` in chat '{res['conversation_name']}':\n{res['new_message']}"
    except Exception as e:
        return f"Error editing Teams message: {e}"


@mcp.tool()
def delete_teams_message(chat_name_or_id: str, message_id: str) -> str:
    """Delete an existing message sent by you in a Teams group chat or 1:1 chat.
    
    Args:
        chat_name_or_id: Name or thread ID of the chat.
        message_id: ID of the message to delete.
    """
    try:
        res = teams_client.delete_message(chat_name_or_id, message_id=message_id)
        return f"✓ Successfully deleted message `{res['message_id']}` from chat '{res['conversation_name']}'."
    except Exception as e:
        return f"Error deleting Teams message: {e}"

if __name__ == "__main__":
    mcp.run()
