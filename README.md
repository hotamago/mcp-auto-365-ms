# MCP Auto 365 MS (Model Context Protocol)

A unified, zero-configuration **Model Context Protocol (MCP)** server enabling AI coding agents (**Claude Code**, **Zed Editor**, **Oh My Pi - OMP**, **Cursor**) to seamlessly interact with **Microsoft 365 (SharePoint, OneDrive, and Microsoft Teams)** directly from their coding harness.

---

## 🌟 Key Highlights

- **Bypasses Microsoft Graph API Restrictions:** Microsoft enforces *Protected APIs* on Teams messages (`Chat.Read`, `ChannelMessage.Read.All`), blocking Azure CLI and standard developer tokens. `mcp-auto-365-ms` automatically retrieves and decrypts your active session tokens (`skypetoken_asm`, `FedAuth`, `rtFa`) directly from Google Chrome on Linux using GNOME Keyring (`org.freedesktop.secrets`).
- **One-Step Team Feed Aggregator:** No more multi-step friction (listing chats then opening each chat one by one). A single call fetches all fresh messages across all active group chats in parallel.
- **Smart Mention Detection with Surrounding Context Window:** Extracts every discussion or task where colleagues or leads tag you (`@Nguyễn Hoàng Sơn`, `@Sơn`, `@all`), **kèm theo ngữ cảnh thảo luận xung quanh (các tin nhắn liền trước và liền sau)** để không bỏ lỡ yêu cầu chi tiết, tài liệu đính kèm, hay chỉ đạo liên quan.
- **Full Lifecycle Teams Messaging (Send, Edit, Delete):** Allows your AI agent to post updates, edit sent messages, or revoke/delete messages across any Teams chat or personal notes.
- **Bidirectional SharePoint Integration (Download, Upload, Versioned Replace):**
  - Download raw files/folders with 100% fidelity without forced markdown degradation.
  - Upload local files directly into SharePoint folders (auto-creates folder hierarchy).
  - Replace existing SharePoint files with automatic version history tracking (`Version 2.0`, `3.0`, etc.) while retaining the exact same sharing links and item IDs.
- **Automated Attachment Downloader:** Seamlessly bridges Teams and SharePoint. Discovers SharePoint/OneDrive links posted inside chat messages and downloads the raw original files directly into your local project.

---

## 🛠️ Complete Tool Catalog (13 Tools)

### 1. Microsoft Teams Automation Tools (9 Tools)
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `get_my_mentions` | **Task & Contextual Mention Tracker:** Extracts all messages mentioning you, **kèm dải tin nhắn ngữ cảnh liền trước và liền sau** để nắm trọn vẹn yêu cầu công việc. | `hours`: Hours back (default: `72`)<br/>`limit`: Max mentions (default: `20`)<br/>`context_before`: Số tin nhắn liền trước (default: `2`)<br/>`context_after`: Số tin nhắn liền sau (default: `2`) |
| `get_recent_team_messages` | **One-Step Feed:** Scans all active group chats in parallel and returns all new messages within a time window. | `hours`: Hours back (default: `48`)<br/>`max_chats`: Max active chats to scan (default: `8`)<br/>`limit_per_chat`: Max messages per chat (default: `8`) |
| `send_teams_message` | **Automated Sender:** Sends a message/notification to any group chat or 1:1 chat when directed. Supports `self`, `me`, `sonnh95@vingroup.net`. | `chat_name_or_id`: Chat name, email or thread ID<br/>`message`: Message text (markdown bold/code supported) |
| `edit_teams_message` | **Message Editor:** Edits an existing message sent by you in a chat. | `chat_name_or_id`: Chat name or thread ID<br/>`message_id`: Message ID to edit<br/>`new_message`: Updated message text |
| `delete_teams_message` | **Message Deleter:** Deletes/revokes a message sent by you in a chat. | `chat_name_or_id`: Chat name or thread ID<br/>`message_id`: Message ID to delete |
| `download_chat_attachments` | **Teams-to-SharePoint Bridge:** Discovers SharePoint links in a chat and downloads the files directly. | `chat_name_or_id`: Chat name or thread ID<br/>`target_dir`: Target directory (default: `docs/sharepoint`) |
| `read_teams_chat` | Reads full discussion history from a specific chat with optional time and mention filtering. Automatically skips deleted messages. | `chat_name_or_id`: Chat name or thread ID<br/>`limit`: Number of messages (default: `30`)<br/>`since`: Date string (`today`, `yesterday`, `YYYY-MM-DD`)<br/>`only_mentions`: Filter to mentions only |
| `search_teams_chat_messages` | Searches across all recent group chats in parallel for topics or keywords. | `query`: Keyword to search (e.g. `'DTC'`, `'S5'`, `'bug'`)<br/>`limit`: Max results (default: `20`) |
| `list_teams_chats` | Lists recent group chats, 1:1 direct chats (including *Chat with yourself*), and meeting threads with metadata. | `limit`: Max chats (default: `30`)<br/>`filter_keyword`: Filter by name |

### 2. SharePoint & OneDrive Tools (4 Tools)
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `read_sharepoint_link` | Explores a folder tree or inspects document metadata & version history. | `url`: SharePoint/OneDrive URL<br/>`max_depth`: Depth for folder traversal (default: `2`) |
| `download_sharepoint_link` | Downloads original binary files or entire folders preserving structure. | `url_or_guid`: File/Folder URL or unique GUID<br/>`target_dir`: Local directory (default: `docs/sharepoint`) |
| `upload_sharepoint_file` | **Uploader:** Uploads a local file to a SharePoint folder (auto-creates folder tree). | `local_file_path`: Path to local file<br/>`target_folder_url_or_path`: SharePoint URL or folder path<br/>`target_file_name`: Optional custom filename |
| `replace_sharepoint_file` | **Versioned Replacer:** Replaces an existing SharePoint file with a new version while preserving the item ID, URL, and recording version history. | `local_file_path`: Path to local updated file<br/>`file_url_or_guid`: File URL or GUID |

---

## 🚀 Quick Start & Installation

### One-Click Installer (Recommended)

Run the included automated setup script:

```bash
git clone git@github.com:hotamago/mcp-auto-365-ms.git
cd mcp-auto-365-ms
chmod +x install.sh
./install.sh
```

---

## 💡 Practical Examples for AI Coding Agents

### Scenario 1: Task Tracking with Context Window
> **User Prompt:** *"Kiểm tra xem gần đây có ai tag tên tôi trong Teams và giao task gì không, đọc cả ngữ cảnh xung quanh để xem cụ thể task là gì."*  
> **Agent Action:** Calls `get_my_mentions(hours=72, context_before=2, context_after=2)`  
> ➔ Hiển thị rõ các tin nhắn liền trước (lý do giao task, link tài liệu) và tin nhắn liền sau (chỉ đạo tiếp theo).

### Scenario 2: Safe Messaging Lifecycle (Send, Edit, Delete)
> **User Prompt:** *"Gửi tin nhắn test vào ghi chú cá nhân của tôi (sonnh95@vingroup.net)"*  
> **Agent Action:** Calls `send_teams_message(...)`, sau đó có thể sửa bằng `edit_teams_message` hoặc thu hồi/xóa bằng `delete_teams_message`.

### Scenario 3: Upload & Version Management on SharePoint
> **User Prompt:** *"Hãy cập nhật file tài liệu `System_Architecture_Design_S5.docx` lên SharePoint."*  
> **Agent Action:** Calls `replace_sharepoint_file(...)`  
> ➔ SharePoint tự động tăng Version 2.0 mà không làm đổi link chia sẻ của dự án.

---

## 🔒 Security & Privacy

- **Local Execution:** All decryptions occur purely in local memory on your machine.
- **Zero Cloud Leakage:** Tokens are never transmitted to any third-party service; requests communicate directly with official Microsoft endpoints (`https://*.sharepoint.com` and `https://*.teams.microsoft.com`).
- **Standard Linux Permissions:** Leverages standard GNOME Secret Service (`org.freedesktop.secrets`) matching your active user desktop session.

---

## 📄 License
MIT License.
