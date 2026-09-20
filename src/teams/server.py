"""Standalone MCP Server for Microsoft Teams chats and messages."""

import sys
from pathlib import Path

# Ensure src root is in python path
src_root = str(Path(__file__).resolve().parent.parent)
if src_root not in sys.path:
    sys.path.insert(0, src_root)

from mcp.server.mcpserver import MCPServer
from teams.client import TeamsClient

mcp = MCPServer("teams-reader")
teams_client = TeamsClient()


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
def read_teams_chat(chat_name_or_id: str, limit: int = 30, since: str = "") -> str:
    """Read full message history and discussions from a specific Teams group chat or 1:1 chat.
    
    Args:
        chat_name_or_id: The exact or partial name of the chat (e.g. 'Proactive Agent', 'ViTa S5', 'Back-end') or the chat thread ID.
        limit: Number of recent messages to fetch (default: 30, max: 100).
        since: Optional filter for messages since a date/time (e.g. '2026-09-18' or '2026-09-18T08:00:00Z').
    """
    try:
        res = teams_client.get_messages(chat_name_or_id, limit=limit, since=since if since else None)
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


if __name__ == "__main__":
    mcp.run()
