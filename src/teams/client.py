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

            space_name = props.get('spaceThreadTopic')
            channel_topic = props.get('topicThreadTopic')

            if c_id == "48:notes":
                display_name = "Chat with yourself (Notes)"
            elif space_name and channel_topic:
                display_name = f"[{space_name}] #{channel_topic}"
            elif topic:
                display_name = topic
            elif chat_type == "DirectChat":
                display_name = f"1:1 Chat ({sender})"
            else:
                display_name = c_id
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

    def get_user_mentions(self, hours: int = 72, limit: int = 20, context_before: int = 2, context_after: int = 2) -> List[Dict[str, Any]]:
        """Search across all active chats for messages specifically mentioning the user,
        including surrounding discussion context (messages before and after the mention).
        """
        convs = self.list_conversations(page_size=25)
        group_chats = [c for c in convs if c["type"] in ["GroupChat", "MeetingChat"]][:15]
        cutoff = datetime.now(timezone.utc) - timedelta(hours=hours)

        def scan_mentions(c):
            found = []
            try:
                fetch_limit = max(30, limit * 2)
                res = self.get_messages(c["id"], limit=fetch_limit)
                msgs = res.get("messages", [])
                for idx, m in enumerate(msgs):
                    try:
                        m_dt = datetime.fromisoformat(m["timestamp"].replace(" ", "T") + "+00:00")
                        if m_dt < cutoff:
                            continue
                    except Exception:
                        pass

                    if any(term in m["content"].lower() for term in ["@nguyễn hoàng sơn", "@hoàng sơn", "@sơn", "@all"]):
                        start_idx = max(0, idx - context_before)
                        end_idx = min(len(msgs), idx + context_after + 1)
                        context_list = []
                        for j in range(start_idx, end_idx):
                            ctx_msg = msgs[j]
                            context_list.append({
                                "sender": ctx_msg["sender"],
                                "timestamp": ctx_msg["timestamp"],
                                "content": ctx_msg["content"],
                                "sharepoint_links": ctx_msg.get("sharepoint_links", []),
                                "is_mention": (j == idx),
                                "offset": j - idx
                            })

                        found.append({
                            "chat_name": c["name"],
                            "chat_id": c["id"],
                            "sender": m["sender"],
                            "timestamp": m["timestamp"],
                            "content": m["content"],
                            "sharepoint_links": m.get("sharepoint_links", []),
                            "context": context_list
                        })
            except Exception:
                pass
            return found

        with ThreadPoolExecutor(max_workers=6) as pool:
            all_mentions = list(pool.map(scan_mentions, group_chats))

        flat_mentions = [item for sublist in all_mentions for item in sublist]
        flat_mentions.sort(key=lambda x: x["timestamp"], reverse=True)
        return flat_mentions[:limit]

    def send_message(self, conversation_id_or_name: str, message: str, reply_to_id: Optional[str] = None, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Send a message to a specific Teams conversation, with optional quote-reply and file attachment."""
        auth = TeamsAuthManager.get_auth()

        conv = self.find_conversation(conversation_id_or_name)
        if not conv:
            raise ValueError(f"Could not resolve conversation '{conversation_id_or_name}'. Please verify the chat name or ID.")

        conv_id = conv["id"]
        conv_name = conv["name"]

        # 1. Handle file attachment if provided
        file_info = None
        if file_path:
            p = Path(file_path).resolve()
            if not p.is_file():
                raise FileNotFoundError(f"Attachment file not found on disk: {file_path}")

            from sharepoint.client import SharePointClient
            sp_client = SharePointClient()
            target_folder = "VF-VSF Collaboration/00.ViTa/04. Squad Sharepoint/S5/02. Technical Docs/Shared from Chat"
            upload_res = sp_client.upload_file(str(p), target_folder_url_or_path=target_folder)
            file_info = upload_res
            sz = upload_res["size"]
            sz_str = f"{sz / (1024*1024):.2f} MB" if sz > 1024*1024 else f"{sz / 1024:.1f} KB"
            message += f"\n\n📎 **Tệp đính kèm:** [{upload_res['name']}]({upload_res['webUrl']}) *({sz_str})*"

        # 2. Convert message text to Teams HTML
        html_content = text_to_teams_html(message)

        # 3. Handle quote reply if reply_to_id is provided
        if reply_to_id:
            quoted_sender = "Member"
            quoted_preview = ""
            try:
                recent = self.get_messages(conv_id, limit=30).get("messages", [])
                for m in recent:
                    if str(m.get("id")) == str(reply_to_id):
                        quoted_sender = m.get("sender", "Member")
                        quoted_preview = m.get("content", "")[:150]
                        break
            except Exception:
                pass

            quote_block = (
                f'<blockquote itemscope itemtype="http://schema.skype.com/Reply" itemid="{reply_to_id}">'
                f'<strong itemprop="mri">{quoted_sender}</strong>'
                f'<span itemprop="time" itemid="{reply_to_id}"></span>'
                f'<p itemprop="preview">{quoted_preview}</p>'
                f'</blockquote>'
            )
            html_content = quote_block + html_content

        encoded_id = urllib.parse.quote(conv_id)
        url = f"{auth['base_url']}/users/ME/conversations/{encoded_id}/messages"

        now_ms = str(int(time.time() * 1000))
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
            "reply_to_id": reply_to_id,
            "attached_file": file_info,
            "server_arrival_time": data.get("OriginalArrivalTime")
        }

    def get_daily_briefing(self, hours: int = 24) -> str:
        """Generate a structured morning executive briefing combining mentions, active chats, and SharePoint activity."""
        mentions = self.get_user_mentions(hours=hours, limit=10, context_before=2, context_after=1)
        feed = self.get_recent_feed(hours=hours, max_chats=6, limit_per_chat=4)

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M")
        out = [
            f"# ☀️ Microsoft 365 Daily Executive Briefing ({now_str})",
            f"*Tổng hợp hoạt động làm việc trong {hours} giờ qua*\n",
            "---"
        ]

        # 1. Direct Mentions & Assigned Tasks
        out.append(f"## 🎯 1. Nhiệm Vụ & Tin Nhắn Tag Tên Bạn ({len(mentions)} mentions)")
        if mentions:
            for m in mentions:
                out.append(f"### 📍 [{m['chat_name']}] — Người tag: **{m['sender']}** ({m['timestamp']})")
                ctx = m.get("context", [])
                if ctx:
                    for c in ctx:
                        time_part = c['timestamp'][11:19] if len(c['timestamp']) >= 19 else c['timestamp']
                        if c["is_mention"]:
                            out.append(f"👉 **[{time_part}] {c['sender']} (MENTION):**")
                            out.append(f"> {c['content']}")
                        else:
                            rel = f"{c['offset']:+d}"
                            out.append(f"- *({rel}) [{time_part}] {c['sender']}:* {c['content']}")
                else:
                    out.append(f"> {m['content']}")
                out.append("")
        else:
            out.append("*(Không có tin nhắn nào tag tên bạn trong khoảng thời gian này)*\n")

        # 2. Key Discussions Across Active Groups
        out.append(f"## 💬 2. Diễn Biến Tại Các Nhóm Đang Thảo Luận ({len(feed)} active chats)")
        if feed:
            for f in feed:
                out.append(f"### 👥 **{f['chat_name']}** *(Hoạt động gần nhất: {f['last_activity'][:19].replace('T', ' ')})*")
                for msg in f.get("messages", [])[-3:]:
                    out.append(f"- **{msg['sender']}**: {msg['content'][:150]}")
                out.append("")
        else:
            out.append("*(Không có thảo luận mới tại các nhóm chat)*\n")

        # 3. SharePoint Recent Documents
        out.append("## 📄 3. Tài Liệu SharePoint Cập Nhật")
        try:
            from sharepoint.client import SharePointClient
            sp = SharePointClient()
            docs = sp.search_files(query="*", max_results=5)
            if docs:
                for d in docs[:5]:
                    sz_str = f"{d['size'] / (1024*1024):.2f} MB" if d['size'] > 1024*1024 else f"{d['size'] / 1024:.1f} KB"
                    out.append(f"- 📄 **[{d['title']}]({d['path']})** *({sz_str}, sửa lúc: {d['modified']} bởi {d['author']})*")
            else:
                out.append("*(Không có tài liệu cập nhật mới)*")
        except Exception:
            out.append("*(Chưa truy xuất tài liệu SharePoint gần đây)*")

        out.append("\n---\n")
        out.append("💡 **Gợi ý hành động:** Bạn có thể dùng `send_teams_message` để trả lời kèm trích dẫn (quote reply) hoặc gửi file đính kèm trực tiếp.")
        return "\n".join(out)

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
        """Search across recent group chats and channels in parallel for messages matching one or multiple keywords."""
        convs = self.list_conversations(page_size=30)
        target_chats = [c for c in convs if c["type"] in ["GroupChat", "MeetingChat", "Channel"]][:15]

        raw_terms = re.split(r'[,;]+', query)
        q_terms = [t.lower().strip() for t in raw_terms if t.strip()]
        if not q_terms:
            q_terms = [query.lower().strip()]

        def search_chat(c):
            matches = []
            try:
                res = self.get_messages(c["id"], limit=30)
                for m in res.get("messages", []):
                    c_low = m["content"].lower()
                    matched_term = next((term for term in q_terms if term in c_low), None)
                    if matched_term:
                        matches.append({
                            "chat_name": c["name"],
                            "chat_id": c["id"],
                            "sender": m["sender"],
                            "timestamp": m["timestamp"],
                            "content": m["content"],
                            "sharepoint_links": m.get("sharepoint_links", []),
                            "matched_keyword": matched_term
                        })
            except Exception:
                pass
            return matches

        with ThreadPoolExecutor(max_workers=6) as pool:
            all_hits = list(pool.map(search_chat, target_chats))

        flat_hits = [item for sublist in all_hits for item in sublist]
        flat_hits.sort(key=lambda x: x["timestamp"], reverse=True)
        return flat_hits[:max_results]
