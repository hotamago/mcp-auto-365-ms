# MCP Auto 365 MS (Model Context Protocol)

A unified, zero-configuration **Model Context Protocol (MCP)** server enabling AI coding agents (**Claude Code**, **Zed Editor**, **Oh My Pi - OMP**, **Cursor**) to seamlessly interact with **Microsoft 365 (SharePoint, OneDrive, and Microsoft Teams)** directly from their coding harness.

---

## 🌟 Key Highlights

- **Bypasses Microsoft Graph API Restrictions:** Microsoft enforces *Protected APIs* on Teams messages (`Chat.Read`, `ChannelMessage.Read.All`), blocking Azure CLI and standard developer tokens. `mcp-auto-365-ms` automatically retrieves and decrypts your active session tokens (`skypetoken_asm`, `FedAuth`, `rtFa`) directly from Google Chrome on Linux using GNOME Keyring (`org.freedesktop.secrets`).
- **No Document Degradation:** Unlike tools that force-convert Office documents into Markdown (losing charts, embedded diagrams, formulas, and formatting), SharePoint downloads deliver original, raw binary files (`.docx`, `.xlsx`, `.pptx`, `.pdf`, `.7z`).
- **Real-Time Group Chat Access:** Agents can inspect discussions, technical decisions, bug reports, and requirements from Teams Group Chats (`@thread.v2`), 1:1 conversations, and meeting threads in under 1 second.
- **Embedded Link Detection:** Automatically detects and parses SharePoint/OneDrive URLs posted inside Teams messages, allowing immediate file downloads.

---

## 🛠️ Registered Tools

### 1. SharePoint & OneDrive Tools
| Tool | Description | Parameters |
| :--- | :--- | :--- |
| `read_sharepoint_link` | Explores a folder tree or inspects document metadata & version history. | `url`: SharePoint/OneDrive URL<br/>`max_depth`: Depth for folder traversal (default: `2`) |
| `download_sharepoint_link` | Downloads original binary files or entire folders preserving structure. | `url_or_guid`: File/Folder URL or unique GUID<br/>`target_dir`: Local directory (default: `docs/sharepoint`) |

### 2. Microsoft Teams Tools
| Tool | Description | Parameters |
| :--- | :--- | :--- |
| `list_teams_chats` | Lists active group chats, direct chats, and meeting threads. | `limit`: Max chats to return (default: `30`)<br/>`filter_keyword`: Filter by chat name or last message |
| `read_teams_chat` | Reads message history, senders, timestamps, and attachments. | `chat_name_or_id`: Partial name (e.g. `'Proactive Agent'`) or Thread ID<br/>`limit`: Max messages (default: `30`)<br/>`since`: ISO date filter (e.g. `'2026-09-18'`) |
| `search_teams_chat_messages` | Searches across all recent group chats for topics/keywords. | `query`: Search term (e.g. `'DTC'`, `'S5'`, `'bug'`, `'review'`)<br/>`limit`: Max results (default: `20`) |

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
   - `mcp-auto-365-ms` (Unified Server — all 5 tools)
   - `mcp-doc-reader` (SharePoint standalone)
   - `mcp-teams-reader` (Teams standalone)
3. Automatically update configurations for **Claude Code**, **Zed**, and **Oh My Pi (OMP)**.

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

### Scenario 1: Tracking requirements from Teams
> **User Prompt:** *"Đọc nhóm chat 'Proactive Agent' từ sáng nay, tóm tắt các phản hồi của team về điều kiện hủy candidate rồi cập nhật vào tài liệu thiết kế."*  
> **Agent Action:** Calls `read_teams_chat(chat_name_or_id="Proactive Agent", since="2026-09-18")`, extracts messages, synthesizes the requirements, and updates `docs/task-09-18.md`.

### Scenario 2: Downloading technical specifications
> **User Prompt:** *"Đồng nghiệp vừa gửi link SharePoint tài liệu kiến trúc SYS3 trong Teams, tải file đó về thư mục docs/sharepoint/ để phân tích."*  
> **Agent Action:** Extracts the URL from the chat message and calls `download_sharepoint_link(url=..., target_dir="docs/sharepoint")`.

### Scenario 3: Investigating reported bugs
> **User Prompt:** *"Tìm trong các nhóm chat xem có ai báo lỗi liên quan đến mã lỗi B1024 hoặc ODO không."*  
> **Agent Action:** Calls `search_teams_chat_messages(query="B1024")` and presents the exact discussion context.

---

## 🔒 Security & Privacy

- **Local Execution:** All decryptions occur purely in local memory on your machine.
- **Zero Cloud Leakage:** Tokens are never transmitted to any third-party service; requests communicate directly with official Microsoft endpoints (`https://*.sharepoint.com` and `https://*.teams.microsoft.com`).
- **Standard Linux Permissions:** Leverages standard GNOME Secret Service (`org.freedesktop.secrets`) matching your active user desktop session.

---

## 📄 License
MIT License.
