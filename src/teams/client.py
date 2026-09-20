"""Microsoft Teams API Client for querying and sending chats and messages."""

import urllib.request
import urllib.parse
import urllib.error
import json
import re
import time
from datetime import datetime, timezone, timedelta
from typing import List, Dict, Any, Optional
from concurrent.futures import ThreadPoolExecutor
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


def text_to_teams_html(text: str) -> str:
    """Convert markdown/plain text to Teams RichText/Html format."""
    paragraphs = text.strip().split('\n\n')
    html_parts = []
    for p in paragraphs:
        # replace single newline with <br/>
        p_html = p.replace('\n', '<br/>')
        # bold **text** -> <b>text</b>
        p_html = re.sub(r'\*\*([^*]+)\*\*', r'<b>\1</b>', p_html)
        # code `code` -> <code>code</code>
        p_html = re.sub(r'`([^`]+)`', r'<code>\1</code>', p_html)
        html_parts.append(f"<p>{p_html}</p>")
    return "".join(html_parts)


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
            if c_id.startswith('48:') and c_id != '48:notes':  # system notification feeds
                continue

            props = c.get('threadProperties', {})
            topic = props.get('topic')
            
            chat_type = "GroupChat"
            if "@thread.v2" in c_id:
                if c_id.startswith("19:meeting_"):
                    chat_type = "MeetingChat"
                else:
                    chat_type = "GroupChat"
            elif "@unq.gbl.spaces" in c_id or c_id == "48:notes":
                chat_type = "DirectChat"
            elif "@thread.tacv2" in c_id:
                chat_type = "Channel"
            last_msg = c.get('lastMessage', {})
            sender = last_msg.get('imdisplayname', 'Unknown')
            raw_content = last_msg.get('content', '')
            clean_msg = clean_teams_html(raw_content)[:120].replace('\n', ' ')
            last_time = last_msg.get('composetime', '')

            if c_id == "48:notes":
                display_name = "Chat with yourself (Notes)"
            else:
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
        ident_lower = identifier.lower().strip()
        if ident_lower in ('48:notes', 'notes', 'self', 'me', 'sonnh95', 'sonnh95@vingroup.net', 'nguyễn hoàng sơn', 'nguyen hoang son'):
            return {
                "id": "48:notes",
                "name": "Chat with yourself (Notes)",
                "type": "DirectChat"
            }

        convs = self.list_conversations(page_size=100)
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

        # 4. Direct ID fallback
        if '@' in identifier or identifier.startswith('48:') or identifier.startswith('19:'):
            return {"id": identifier, "name": identifier, "type": "DirectChat"}

        return None

    def get_messages(self, conversation_id_or_name: str, limit: int = 30, since: Optional[str] = None, only_mentions: bool = False) -> Dict[str, Any]:
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
                now_utc = datetime.now(timezone.utc)
                if since.lower() == "today":
                    since_dt = now_utc.replace(hour=0, minute=0, second=0, microsecond=0)
                elif since.lower() == "yesterday":
                    since_dt = (now_utc - timedelta(days=1)).replace(hour=0, minute=0, second=0, microsecond=0)
                elif len(since) == 10:  # YYYY-MM-DD
                    since_dt = datetime.fromisoformat(f"{since}T00:00:00+00:00")
                else:
                    since_dt = datetime.fromisoformat(since.replace('Z', '+00:00'))
            except Exception:
                pass

        for m in reversed(raw_messages):
            m_type = m.get('messagetype')
            if m_type not in ['Text', 'RichText/Html']:
                continue

            # Skip deleted messages
            if m.get('properties', {}).get('deletetime'):
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

            if only_mentions:
                if not any(term in cleaned_text.lower() for term in ['@nguyễn hoàng sơn', '@hoàng sơn', '@sơn', '@all']):
                    continue

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

    def get_recent_feed(self, hours: int = 48, max_chats: int = 8, limit_per_chat: int = 8, filter_keyword: str = "") -> List[Dict[str, Any]]:
        """Fetch new messages across all recently active group/meeting chats in parallel."""
        convs = self.list_conversations(page_size=20, filter_keyword=filter_keyword)
        group_chats = [c for c in convs if c["type"] in ["GroupChat", "MeetingChat"]][:max_chats]

        cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)

        def fetch_chat(c):
            try:
                res = self.get_messages(c["id"], limit=limit_per_chat)
                recent_msgs = []
                for m in res.get("messages", []):
                    try:
                        m_dt = datetime.fromisoformat(m["timestamp"].replace(" ", "T") + "+00:00")
                        if m_dt >= cutoff:
                            recent_msgs.append(m)
                    except Exception:
                        recent_msgs.append(m)
                if recent_msgs:
                    return {
                        "chat_name": c["name"],
                        "chat_id": c["id"],
                        "last_activity": c["last_activity"],
                        "messages": recent_msgs
                    }
            except Exception:
                pass
            return None

        with ThreadPoolExecutor(max_workers=6) as pool:
            results = list(pool.map(fetch_chat, group_chats))

        return [r for r in results if r is not None]

    def get_user_mentions(self, hours: int = 72, limit: int = 20) -> List[Dict[str, Any]]:
        """Search across all active chats for messages specifically mentioning the user."""
        convs = self.list_conversations(page_size=20)
        group_chats = [c for c in convs if c["type"] in ["GroupChat", "MeetingChat"]][:12]
        cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)

        def scan_mentions(c):
            found = []
            try:
                res = self.get_messages(c["id"], limit=20)
                for m in res.get("messages", []):
                    try:
                        m_dt = datetime.fromisoformat(m["timestamp"].replace(" ", "T") + "+00:00")
                        if m_dt < cutoff:
                            continue
                    except Exception:
                        pass

                    if any(term in m["content"].lower() for term in ["@nguyễn hoàng sơn", "@hoàng sơn", "@sơn", "@all"]):
                        found.append({
                            "chat_name": c["name"],
                            "chat_id": c["id"],
                            "sender": m["sender"],
                            "timestamp": m["timestamp"],
                            "content": m["content"],
                            "sharepoint_links": m.get("sharepoint_links", [])
                        })
            except Exception:
                pass
            return found

        with ThreadPoolExecutor(max_workers=6) as pool:
            all_mentions = list(pool.map(scan_mentions, group_chats))

        flat_mentions = [item for sublist in all_mentions for item in sublist]
        flat_mentions.sort(key=lambda x: x["timestamp"], reverse=True)
        return flat_mentions[:limit]

    def send_message(self, conversation_id_or_name: str, message: str) -> Dict[str, Any]:
        """Send a message to a specific Teams conversation."""
        auth = TeamsAuthManager.get_auth()

        conv = self.find_conversation(conversation_id_or_name)
        if not conv:
            raise ValueError(f"Could not resolve conversation '{conversation_id_or_name}'. Please verify the chat name or ID.")

        conv_id = conv["id"]
        conv_name = conv["name"]

        encoded_id = urllib.parse.quote(conv_id)
        url = f"{auth['base_url']}/users/ME/conversations/{encoded_id}/messages"

        now_ms = str(int(time.time() * 1000))
        html_content = text_to_teams_html(message)

        display_name = auth.get("claims", {}).get("name", "Nguyễn Hoàng Sơn (VF-KPTX-VPTAITX)")

        payload = {
            "content": html_content,
            "messagetype": "RichText/Html",
            "contenttype": "text",
            "clientmessageid": now_ms,
            "imdisplayname": display_name
        }

        headers = self._get_headers()
        headers["Content-Type"] = "application/json"

        req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method='POST')
        with urllib.request.urlopen(req) as resp:
            data = json.loads(resp.read().decode('utf-8'))

        return {
            "status": "SENT",
            "conversation_id": conv_id,
            "conversation_name": conv_name,
            "message_sent": message,
            "server_arrival_time": data.get("OriginalArrivalTime")
        }

    def edit_message(self, conversation_id_or_name: str, message_id: str, new_message: str) -> Dict[str, Any]:
        """Edit an existing message in a Teams conversation."""
        auth = TeamsAuthManager.get_auth()

        conv = self.find_conversation(conversation_id_or_name)
        if not conv:
            raise ValueError(f"Could not resolve conversation '{conversation_id_or_name}'.")

        conv_id = conv["id"]
        conv_name = conv["name"]

        encoded_id = urllib.parse.quote(conv_id)
        encoded_msg_id = urllib.parse.quote(str(message_id))
        url = f"{auth['base_url']}/users/ME/conversations/{encoded_id}/messages/{encoded_msg_id}"

        html_content = text_to_teams_html(new_message)

        payload = {
            "content": html_content,
            "messagetype": "RichText/Html",
            "contenttype": "text"
        }

        headers = self._get_headers()
        headers["Content-Type"] = "application/json"

        req = urllib.request.Request(url, data=json.dumps(payload).encode('utf-8'), headers=headers, method='PUT')
        with urllib.request.urlopen(req) as resp:
            pass

        return {
            "status": "EDITED",
            "conversation_id": conv_id,
            "conversation_name": conv_name,
            "message_id": message_id,
            "new_message": new_message
        }

    def delete_message(self, conversation_id_or_name: str, message_id: str) -> Dict[str, Any]:
        """Delete an existing message in a Teams conversation."""
        auth = TeamsAuthManager.get_auth()

        conv = self.find_conversation(conversation_id_or_name)
        if not conv:
            raise ValueError(f"Could not resolve conversation '{conversation_id_or_name}'.")

        conv_id = conv["id"]
        conv_name = conv["name"]

        encoded_id = urllib.parse.quote(conv_id)
        encoded_msg_id = urllib.parse.quote(str(message_id))
        url = f"{auth['base_url']}/users/ME/conversations/{encoded_id}/messages/{encoded_msg_id}"

        headers = self._get_headers()

        req = urllib.request.Request(url, headers=headers, method='DELETE')
        with urllib.request.urlopen(req) as resp:
            pass

        return {
            "status": "DELETED",
            "conversation_id": conv_id,
            "conversation_name": conv_name,
            "message_id": message_id
        }

    def search_messages(self, query: str, max_results: int = 20) -> List[Dict[str, Any]]:
        """Search across recent group chats in parallel for messages containing query."""
        convs = self.list_conversations(page_size=25)
        target_chats = [c for c in convs if c["type"] in ["GroupChat", "MeetingChat"]][:12]
        q_lower = query.lower().strip()

        def search_chat(c):
            matches = []
            try:
                res = self.get_messages(c["id"], limit=25)
                for m in res.get("messages", []):
                    if q_lower in m["content"].lower():
                        matches.append({
                            "chat_name": c["name"],
                            "chat_id": c["id"],
                            "sender": m["sender"],
                            "timestamp": m["timestamp"],
                            "content": m["content"],
                            "sharepoint_links": m.get("sharepoint_links", [])
                        })
            except Exception:
                pass
            return matches

        with ThreadPoolExecutor(max_workers=6) as pool:
            all_hits = list(pool.map(search_chat, target_chats))

        flat_hits = [item for sublist in all_hits for item in sublist]
        flat_hits.sort(key=lambda x: x["timestamp"], reverse=True)
        return flat_hits[:max_results]
