# MCP Auto 365 MS (Model Context Protocol)

A unified, zero-configuration **Model Context Protocol (MCP)** server enabling AI coding agents (**Claude Code**, **Zed Editor**, **Oh My Pi - OMP**, **Cursor**) to seamlessly interact with **Microsoft 365 (SharePoint, OneDrive, and Microsoft Teams)** directly from their coding harness.

---

## 🌟 Key Highlights

- **Bypasses Microsoft Graph API Restrictions:** Microsoft enforces *Protected APIs* on Teams messages (`Chat.Read`, `ChannelMessage.Read.All`), blocking Azure CLI and standard developer tokens. `mcp-auto-365-ms` automatically retrieves and decrypts your active session tokens (`skypetoken_asm`, `FedAuth`, `rtFa`) directly from Google Chrome on Linux using GNOME Keyring (`org.freedesktop.secrets`).
- **One-Step Team Feed Aggregator:** No more multi-step friction (listing chats then opening each chat one by one). A single call fetches all fresh messages across all active group chats in parallel.
- **Smart Mention Detection:** Instantly extracts every discussion or task where colleagues or leads tag you (`@Nguyễn Hoàng Sơn`, `@Sơn`, `@all`).
- **Bidirectional Teams Messaging:** Allows your AI agent to automatically post updates, test results, bug summaries, or replies to any Teams chat upon your instruction.
- **Automated Attachment Downloader:** Seamlessly bridges Teams and SharePoint. Discovers SharePoint/OneDrive links posted inside chat messages and downloads the raw original files with 100% fidelity directly to your local project.
- **No Document Degradation:** Downloads original binary files (`.docx`, `.xlsx`, `.pptx`, `.pdf`, `.7z`) without forced Markdown conversion.

---

## 🛠️ Complete Tool Catalog (9 Tools)

### 1. Microsoft Teams Automation Tools
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `get_recent_team_messages` | **One-Step Feed:** Scans all active group chats in parallel and returns all new messages within a time window. | `hours`: Hours back (default: `48`)<br/>`max_chats`: Max active chats to scan (default: `8`)<br/>`limit_per_chat`: Max messages per chat (default: `8`) |
| `get_my_mentions` | **Task & Mention Tracker:** Extracts all messages across all chats mentioning you. | `hours`: Hours back (default: `72`)<br/>`limit`: Max mentions to return (default: `20`) |
| `send_teams_message` | **Automated Sender:** Sends a message/notification to any group chat or 1:1 chat when directed. | `chat_name_or_id`: Chat name or thread ID<br/>`message`: Message text (markdown bold/code supported) |
| `download_chat_attachments` | **Teams-to-SharePoint Bridge:** Discovers SharePoint links in a chat and downloads the files directly. | `chat_name_or_id`: Chat name or thread ID<br/>`target_dir`: Target directory (default: `docs/sharepoint`) |
| `read_teams_chat` | Reads full discussion history from a specific chat with optional time and mention filtering. | `chat_name_or_id`: Chat name or thread ID<br/>`limit`: Number of messages (default: `30`)<br/>`since`: Date string (`today`, `yesterday`, `YYYY-MM-DD`)<br/>`only_mentions`: Filter to mentions only |
| `search_teams_chat_messages` | Searches across all recent group chats in parallel for topics or keywords. | `query`: Keyword to search (e.g. `'DTC'`, `'S5'`, `'bug'`)<br/>`limit`: Max results (default: `20`) |
| `list_teams_chats` | Lists recent group chats, 1:1 direct chats, and meeting threads with metadata. | `limit`: Max chats (default: `30`)<br/>`filter_keyword`: Filter by name |

### 2. SharePoint & OneDrive Tools
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `read_sharepoint_link` | Explores a folder tree or inspects document metadata & version history. | `url`: SharePoint/OneDrive URL<br/>`max_depth`: Depth for folder traversal (default: `2`) |
| `download_sharepoint_link` | Downloads original binary files or entire folders preserving structure. | `url_or_guid`: File/Folder URL or unique GUID<br/>`target_dir`: Local directory (default: `docs/sharepoint`) |

---

## 🚀 Quick Start & Installation

### Option 1: One-Click Installer (Recommended)

Run the included automated setup script:

```bash
git clone git@github.com:hotamago/mcp-auto-365-ms.git
cd mcp-auto-365-ms
chmod +x install.sh
./install.sh
```

The script will:
1. Verify Python 3 and install required dependencies (`mcp`, `cryptography`, `dbus-python`).
2. Create symlinks in `~/.local/bin/`:
   - `mcp-auto-365-ms` (Unified Server — all 9 tools)
   - `mcp-doc-reader` (SharePoint standalone)
   - `mcp-teams-reader` (Teams standalone)
3. Automatically update configurations for **Claude Code**, **Zed Editor**, and **Oh My Pi (OMP)**.

---

### Option 2: Manual Configuration

#### 1. Claude Code (`~/.claude.json`)
```json
{
  "mcpServers": {
    "auto-365-ms": {
      "command": "/home/sonnh95/.local/bin/mcp-auto-365-ms"
    }
  }
}
```

#### 2. Zed Editor (`~/.config/zed/settings.json`)
```json
{
  "context_servers": {
    "auto-365-ms": {
      "command": {
        "path": "/home/sonnh95/.local/bin/mcp-auto-365-ms",
        "args": []
      }
    }
  }
}
```

#### 3. Oh My Pi (`~/.omp/agent/mcp.json` or `.omp/mcp.json`)
```json
{
  "$schema": "https://raw.githubusercontent.com/can1357/oh-my-pi/main/packages/coding-agent/src/config/mcp-schema.json",
  "mcpServers": {
    "auto-365-ms": {
      "command": "/home/sonnh95/.local/bin/mcp-auto-365-ms"
    }
  }
}
```

---

## 💡 Practical Examples for AI Coding Agents

### Scenario 1: One-Step Morning Catch-Up
> **User Prompt:** *"Hôm nay team thảo luận những gì trong các nhóm chat? Có yêu cầu gì mới không?"*  
> **Agent Action:** Calls `get_recent_team_messages(hours=24)`, aggregates discussions from `[VF - Internal - ViTa] Proactive Agent Feature Dev`, `[S6 - Vita] Back-end x Infra`, etc., and gives a structured briefing.

### Scenario 2: What tasks did colleagues assign to me?
> **User Prompt:** *"Kiểm tra xem gần đây có ai tag tên tôi trong Teams và giao task gì không."*  
> **Agent Action:** Calls `get_my_mentions(hours=72)` and immediately reports:
> - Nguyễn Phú Hiển assigned Feature 1 (`r811`) to you.
> - Nguyễn Việt Phương mentioned CAN signal transmission guidance for you.

### Scenario 3: Automated document download from chat
> **User Prompt:** *"Đồng nghiệp vừa gửi file kiến trúc trong nhóm Proactive Agent, tải về thư mục docs/sharepoint/."*  
> **Agent Action:** Calls `download_chat_attachments(chat_name_or_id="Proactive Agent")`, auto-detects the SharePoint URL, and downloads the binary document directly.

### Scenario 4: Automated Notification & Reporting
> **User Prompt:** *"Báo cáo vào nhóm 'Proactive Agent' rằng tôi đã hoàn thành việc rà soát tín hiệu CVC/SIDL và cập nhật file task-09-18.md."*  
> **Agent Action:** Calls `send_teams_message(chat_name_or_id="Proactive Agent", message="Đã hoàn thành rà soát tín hiệu...")`.

---

## 🔒 Security & Privacy

- **Local Execution:** All decryptions occur purely in local memory on your machine.
- **Zero Cloud Leakage:** Tokens are never transmitted to any third-party service; requests communicate directly with official Microsoft endpoints (`https://*.sharepoint.com` and `https://*.teams.microsoft.com`).
- **Standard Linux Permissions:** Leverages standard GNOME Secret Service (`org.freedesktop.secrets`) matching your active user desktop session.

---

## 📄 License
MIT License.
