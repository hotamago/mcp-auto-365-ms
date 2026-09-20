"""Microsoft Teams API Client for querying chats and messages."""

import urllib.request
import urllib.parse
import urllib.error
import json
import re
from datetime import datetime
from typing import List, Dict, Any, Optional
from .auth import TeamsAuthManager


def clean_teams_html(html_content: str) -> str:
    """Convert Teams HTML messages into readable markdown/plain text."""
    if not html_content:
        return ""

    text = html_content
    # Replace line breaks
    text = re.sub(r'<br\s*/?>', '\n', text)
    text = re.sub(r'</p>', '\n', text)
    text = re.sub(r'<p>', '', text)

    # Extract mentions: <span itemtype=".../Mention">Name</span>
    text = re.sub(r'<span[^>]*itemtype="[^"]*Mention"[^>]*>([^<]+)</span>', r'@\1', text)

    # Extract links: <a href="...">Text</a> -> [Text](href)
    text = re.sub(r'<a\s+[^>]*href="([^"]+)"[^>]*>([^<]+)</a>', r'[\2](\1)', text)

    # Clean other tags
    text = re.sub(r'<[^>]+>', '', text)

    # Decode HTML entities
    text = text.replace('&nbsp;', ' ')
    text = text.replace('&quot;', '"')
    text = text.replace('&amp;', '&')
    text = text.replace('&lt;', '<')
    text = text.replace('&gt;', '>')

    # Normalize multiple newlines
    text = re.sub(r'\n{3,}', '\n\n', text)
    return text.strip()


class TeamsClient:
    def __init__(self):
        self._auth = TeamsAuthManager.get_auth()

    def _get_headers(self) -> Dict[str, str]:
        auth = TeamsAuthManager.get_auth()
        return {
            'Authentication': f"skypetoken={auth['token']}",
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36',
            'Accept': 'application/json',
            'x-ms-client-version': '27/24081200000',
            'x-ms-client-env': 'prod',
            'Origin': 'https://teams.microsoft.com',
            'Referer': 'https://teams.microsoft.com/'
        }

    def list_conversations(self, page_size: int = 50, filter_keyword: str = "") -> List[Dict[str, Any]]:
        """List active group chats, direct chats, and meeting threads."""
        auth = TeamsAuthManager.get_auth()
        url = f"{auth['base_url']}/users/ME/conversations?view=msnp24Equivalent&pageSize={page_size}"
        
        req = urllib.request.Request(url, headers=self._get_headers())
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))

        convs = data.get('conversations', [])
        results = []

        kw = filter_keyword.lower().strip() if filter_keyword else ""

        for c in convs:
            c_id = c.get('id', '')
            if c_id.startswith('48:'):  # system notification feeds
                continue

            props = c.get('threadProperties', {})
            topic = props.get('topic')
            
            chat_type = "GroupChat"
            if "@thread.v2" in c_id:
                if c_id.startswith("19:meeting_"):
                    chat_type = "MeetingChat"
                else:
                    chat_type = "GroupChat"
            elif "@unq.gbl.spaces" in c_id:
                chat_type = "DirectChat"
            elif "@thread.tacv2" in c_id:
                chat_type = "Channel"

            last_msg = c.get('lastMessage', {})
            sender = last_msg.get('imdisplayname', 'Unknown')
            raw_content = last_msg.get('content', '')
            clean_msg = clean_teams_html(raw_content)[:120].replace('\n', ' ')
            last_time = last_msg.get('composetime', '')

            display_name = topic or (f"1:1 Chat ({sender})" if chat_type == "DirectChat" else c_id)

            item = {
                "id": c_id,
                "name": display_name,
                "type": chat_type,
                "last_activity": last_time,
                "last_sender": sender,
                "last_message": clean_msg
            }

            if kw:
                if kw not in display_name.lower() and kw not in clean_msg.lower():
                    continue

            results.append(item)

        return results

    def find_conversation(self, identifier: str) -> Optional[Dict[str, Any]]:
        """Find a conversation by exact ID or fuzzy title match."""
        convs = self.list_conversations(page_size=100)
        ident_lower = identifier.lower().strip()

        # 1. Exact ID match
        for c in convs:
            if c["id"] == identifier:
                return c

        # 2. Exact Title match
        for c in convs:
            if c["name"].lower() == ident_lower:
                return c

        # 3. Substring match
        for c in convs:
            if ident_lower in c["name"].lower():
                return c

        return None

    def get_messages(self, conversation_id_or_name: str, limit: int = 30, since: Optional[str] = None) -> Dict[str, Any]:
        """Retrieve recent messages for a conversation."""
        auth = TeamsAuthManager.get_auth()

        conv = self.find_conversation(conversation_id_or_name)
        conv_id = conv["id"] if conv else conversation_id_or_name
        conv_name = conv["name"] if conv else conv_id

        encoded_id = urllib.parse.quote(conv_id)
        url = f"{auth['base_url']}/users/ME/conversations/{encoded_id}/messages?pageSize={limit}"

        req = urllib.request.Request(url, headers=self._get_headers())
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))

        raw_messages = data.get('messages', [])
        formatted = []

        since_dt = None
        if since:
            try:
                if len(since) == 10:  # YYYY-MM-DD
                    since_dt = datetime.fromisoformat(f"{since}T00:00:00Z")
                else:
                    since_dt = datetime.fromisoformat(since.replace('Z', '+00:00'))
            except Exception:
                pass

        for m in reversed(raw_messages):
            m_type = m.get('messagetype')
            if m_type not in ['Text', 'RichText/Html']:
                continue

            comp_time_str = m.get('composetime', '')
            if since_dt and comp_time_str:
                try:
                    msg_dt = datetime.fromisoformat(comp_time_str.replace('Z', '+00:00'))
                    if msg_dt < since_dt:
                        continue
                except Exception:
                    pass

            sender = m.get('imdisplayname', 'Unknown')
            raw_content = m.get('content', '')
            cleaned_text = clean_teams_html(raw_content)

            # Extract attachments / sharepoint links
            sp_links = re.findall(r'https://[a-zA-Z0-9_-]*sharepoint\.com[^\s"\'<>]+', raw_content)

            formatted.append({
                "id": m.get('id'),
                "sender": sender,
                "timestamp": comp_time_str[:19].replace('T', ' '),
                "content": cleaned_text,
                "sharepoint_links": sp_links
            })

        return {
            "conversation_id": conv_id,
            "conversation_name": conv_name,
            "total_messages": len(formatted),
            "messages": formatted
        }

    def search_messages(self, query: str, max_results: int = 20) -> List[Dict[str, Any]]:
        """Search across recent group chats for messages containing the query."""
        convs = self.list_conversations(page_size=30)
        q_lower = query.lower().strip()
        hits = []

        for c in convs:
            if c["type"] not in ["GroupChat", "MeetingChat"]:
                continue

            try:
                res = self.get_messages(c["id"], limit=20)
                for m in res.get("messages", []):
                    if q_lower in m["content"].lower():
                        hits.append({
                            "chat_name": c["name"],
                            "chat_id": c["id"],
                            "sender": m["sender"],
                            "timestamp": m["timestamp"],
                            "content": m["content"],
                            "sharepoint_links": m.get("sharepoint_links", [])
                        })
                        if len(hits) >= max_results:
                            return hits
            except Exception:
                continue

        return hits
