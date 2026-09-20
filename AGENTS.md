# 🤖 AGENTS.md — Developer & AI Agent Operations Manual

> **Target Audience:** AI Coding Agents (Claude Code, Oh My Pi, Zed Editor, Cursor) and Software Engineers maintaining or extending `mcp-auto-365-ms`.
> **Core Mission:** Provide autonomous, zero-friction access to Microsoft 365 (SharePoint, OneDrive, and Microsoft Teams) directly from coding harnesses without administrative Graph API permissions or document quality degradation.

---

## 📌 Table of Contents
1. [Project Overview & Design Philosophy](#1-project-overview--design-philosophy)
2. [Repository Layout & Module Structure](#2-repository-layout--module-structure)
3. [Authentication Invariants & Reverse Engineering](#3-authentication-invariants--reverse-engineering)
4. [Tool Inventory & Interface Contracts (16 Tools)](#4-tool-inventory--interface-contracts-16-tools)
5. [Agent Operational Guidelines & Safety Rules](#5-agent-operational-guidelines--safety-rules)
6. [Testing & Verification Procedures](#6-testing--verification-procedures)
7. [Deployment & Multi-Harness Configuration](#7-deployment--multi-harness-configuration)
8. [Known Pitfalls & Troubleshooting Matrix](#8-known-pitfalls--troubleshooting-matrix)

---

## 1. Project Overview & Design Philosophy

`mcp-auto-365-ms` is a unified **Model Context Protocol (MCP)** server built in Python on top of the `mcp` SDK (FastMCP architecture).

### Core Problems Solved:
1. **Bypassing Microsoft Graph Protected APIs:** Microsoft blocks developer and CLI tokens from reading Teams messages (`Chat.Read`, `ChannelMessage.Read.All` return HTTP 403 Forbidden without tenant admin approval). This project extracts active session tokens directly from Google Chrome's local storage.
2. **Zero Office Document Degradation:** Unlike legacy MCP servers that force convert `.docx`, `.xlsx`, and `.pptx` into lossy Markdown, this project handles raw binaries, preserving tables, formulas, formatting, and vector diagrams with 100% fidelity.
3. **Autonomous, Single-Step Workflows:** AI agents should not make 3 sequential calls when 1 call suffices. Complex operations (e.g. parallel multi-chat feed aggregation, contextual mentions with surrounding message windows, quote-reply with automatic SharePoint file attachments) are unified into single-step tools.
4. **Single-Process Unified Architecture:** Consolidates all 16 tools into a single daemon process (`auto-365-ms`), preventing process bloat and memory leaks across IDEs.

---

## 2. Repository Layout & Module Structure

```text
mcp-auto-365-ms/
├── bin/                              # Symlink-resolving executable launchers
│   ├── mcp-auto-365-ms               # Primary unified server launcher (all 16 tools)
│   ├── mcp-doc-reader                # Standalone SharePoint server launcher
│   └── mcp-teams-reader              # Standalone Teams server launcher
├── src/
│   ├── __init__.py
│   ├── server.py                     # Unified FastMCP server registering all 16 tools
│   ├── common/
│   │   ├── __init__.py
│   │   └── chrome_cookies.py         # GNOME Keyring & SQLite cookie decrypter
│   ├── sharepoint/
│   │   ├── __init__.py
│   │   ├── client.py                 # SharePoint REST & Graph client (search, download, upload, diff)
│   │   └── server.py                 # Standalone FastMCP server for SharePoint
│   └── teams/
│       ├── __init__.py
│       ├── auth.py                   # Teams session extractor (skypetoken_asm & region)
│       ├── client.py                 # Teams Chat Service client (feed, mentions, send, edit, del)
│       └── server.py                 # Standalone FastMCP server for Teams
├── install.sh                        # 1-click automated multi-harness installer
├── requirements.txt                  # Python runtime dependencies
├── README.md                         # End-user documentation
└── AGENTS.md                         # This architecture and operations manual
```

---

## 3. Authentication Invariants & Reverse Engineering

### 3.1 Google Chrome Cookie Decryption (Linux / GNOME Keyring)
- **Path:** `~/.config/google-chrome/Default/Cookies` (SQLite database).
- **Master Key Acquisition:** Queried via DBus from `org.freedesktop.secrets` (Collection `/login/2`).
- **Key Derivation:** `PBKDF2-HMAC-SHA1(password, salt=b'saltysalt', iterations=1, length=16)`.
- **Value Decryption:** AES-128-CBC (`iv = b' ' * 16`). Prefix `v10`/`v11` is stripped, decrypted, and unpadded; the first 32 bytes contain the hash prefix and are trimmed.

### 3.2 Microsoft Teams Authentication
- **Token:** `skypetoken_asm` cookie extracted from `teams.microsoft.com`.
- **Chat Service Regional Base URL:** e.g. `https://apac.ng.msg.teams.microsoft.com/v1`.
- **Request Headers:**
  ```http
  Authentication: skypetoken=<token>
  Accept: application/json
  Content-Type: application/json
  ```
- **User Object ID / Skype ID:** Formatted as `8:orgid:<Azure-AD-GUID>` (e.g. `8:orgid:b6cf511d-9f31-4a84-89d8-3a400a1a544f`).
- **Personal Notes (Chat with yourself):** The dedicated conversation identifier is `48:notes`.
- **Team Channels:** Identifiers end in `@thread.tacv2`. The channel display name must be composed from `spaceThreadTopic` and `topicThreadTopic` (`[{space}] #{topic}`).

### 3.3 SharePoint & OneDrive Authentication (Dual-Channel Architecture)
- **Channel 1 (Graph API via Azure CLI):**
  - Token retrieved via `az account get-access-token --resource https://graph.microsoft.com`.
  - Used for folder creation (`POST /children`), direct content upload (`PUT /content`), and replace operations.
  - Automatically handles 401 expiration challenges by clearing the cached token and re-invoking `az`.
- **Channel 2 (Direct Session Cookies):**
  - Uses `FedAuth` and `rtFa` cookies decrypted from Chrome.
  - **Header Requirement:** `Cookie: rtFa=...; FedAuth=...` with User-Agent `Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36`.
  - Used for direct document downloads (HTTP 200 bypassing Graph 403), SharePoint REST search (`/_api/search/query`), and historical version downloads (`/_vti_history/...`).

---

## 4. Tool Inventory & Interface Contracts (16 Tools)

All tools are exposed over JSON-RPC stdio by `src/server.py`.

### 4.1 SharePoint & OneDrive Suite (6 Tools)
1. **`search_sharepoint_files(query: str, max_results: int = 20, file_extension: str = "") -> str`**
   - Fast full-text document search via SharePoint REST API (`~1s`).
   - Returns Document Title, Path, Size, Last Modified Date, Author, and `UniqueId`.
2. **`compare_sharepoint_versions(file_a: str, file_b: str = "", version_a: str = "", version_b: str = "") -> str`**
   - Smart Diff tool.
   - If `file_b` is given: Compares `file_a` against `file_b` (supports local file vs SharePoint file, or two SharePoint files).
   - If `file_b` is omitted: Compares version A vs version B of a single SharePoint file.
   - Outputs unified diff (`+`, `-`) for `.docx` (extracted paragraphs), text/code, and sheet listings for `.xlsx`.
3. **`read_sharepoint_link(url: str, max_depth: int = 2) -> str`**
   - If folder: Recursively traverses directory tree, child counts, authors, sizes.
   - If document: Shows metadata, author, modified date, and version history.
4. **`download_sharepoint_link(url_or_guid: str, target_dir: str = "docs/sharepoint") -> str`**
   - Downloads original binary files or entire folders with 100% fidelity without Markdown conversion.
5. **`upload_sharepoint_file(local_file_path: str, target_folder_url_or_path: str, target_file_name: str = "") -> str`**
   - Uploads a local file to SharePoint, automatically generating all missing parent folders.
6. **`replace_sharepoint_file(local_file_path: str, file_url_or_guid: str) -> str`**
   - Replaces an existing file in-place, generating a new version (`Version 2.0`) while preserving sharing links and Item ID.

### 4.2 Microsoft Teams Suite (10 Tools)
7. **`get_daily_briefing(hours: int = 24) -> str`**
   - One-click executive morning briefing aggregating:
     1. Tasks & direct mentions (with context).
     2. Active team discussions.
     3. Recently updated SharePoint documents.
8. **`get_my_mentions(hours: int = 72, limit: int = 20, context_before: int = 2, context_after: int = 2) -> str`**
   - Scans active group chats in parallel for mentions (`@Nguyễn Hoàng Sơn`, `@Sơn`, `@all`).
   - Returns a context window of preceding (`context_before`) and following (`context_after`) messages to preserve full task background.
9. **`get_recent_team_messages(hours: int = 48, max_chats: int = 8, limit_per_chat: int = 8, filter_keyword: str = "") -> str`**
   - 1-step team feed aggregator scanning active chats concurrently via `ThreadPoolExecutor(max_workers=6)` (`< 3s`).
10. **`send_teams_message(chat_name_or_id: str, message: str, reply_to_id: str = "", file_path: str = "") -> str`**
    - Multi-purpose sender:
      - Plain text / Markdown message.
      - **Quote-Reply:** Prepends standard Teams `<blockquote ...>` referencing `reply_to_id`.
      - **Attachment:** Uploads `file_path` to SharePoint and embeds attachment card/link.
11. **`edit_teams_message(chat_name_or_id: str, message_id: str, new_message: str) -> str`**
    - Edits an already sent message via `PUT .../messages/{id}`.
12. **`delete_teams_message(chat_name_or_id: str, message_id: str) -> str`**
    - Revokes/deletes a message via `DELETE .../messages/{id}`.
13. **`download_chat_attachments(chat_name_or_id: str, target_dir: str = "docs/sharepoint", limit: int = 5) -> str`**
    - Auto-detects SharePoint URLs in chat messages and downloads the raw files directly.
14. **`read_teams_chat(chat_name_or_id: str, limit: int = 30, since: str = "", only_mentions: bool = False, output_file: str = "") -> str`**
    - Reads chat history (automatically filtering deleted messages).
    - If `output_file` is specified, saves formatted chat transcript directly to Markdown on disk.
15. **`list_teams_chats(limit: int = 30, filter_keyword: str = "") -> str`**
    - Lists group chats, 1:1 direct chats, and channels (`[Team] #Channel`).
16. **`search_teams_chat_messages(keywords: list[str], limit: int = 20) -> str`**
    - Concurrent keyword search across group chats and channels using an explicit array schema.

---

## 5. Agent Operational Guidelines & Safety Rules

<critical>
### ⚠️ STRICT SAFETY DIRECTIVES:
1. **TESTING DESTINATIONS ONLY:** When testing `send_teams_message`, `edit_teams_message`, or `delete_teams_message`, you MUST ONLY target the user's personal self-chat (`sonnh95@vingroup.net`, `self`, `me`, or `48:notes`). NEVER send, edit, or delete messages in team group chats, channels, or direct chats with colleagues!
2. **CLEANUP AFTER TESTING:** Any test messages sent to `48:notes` during automated verification MUST be cleaned up via `delete_message` before completing your turn.
3. **IN-MEMORY PROCESS RELOADING:** Modifying Python code files does NOT immediately reload the running MCP daemon in OMP or IDE harnesses. When modifying server files (`src/server.py`, `src/teams/client.py`, `src/sharepoint/client.py`), you MUST gracefully terminate the existing daemon (`pkill -f "python3.*mcp-auto-365-ms/src/server.py"`) so the harness respawns the fresh code.
4. **NO DUPLICATE SERVER REGISTRATIONS:** Maintain ONLY `auto-365-ms` in `~/.omp/agent/mcp.json`, `~/.claude.json`, and `~/.config/zed/settings.json`. NEVER register `doc-reader` or `teams-reader` concurrently, as they duplicate all tools.
</critical>

---

## 6. Testing & Verification Procedures

Always verify tools using Python JSON-RPC stdio rather than guessing:

```bash
# 1. Verify syntax
python3 -m py_compile src/server.py src/teams/client.py src/sharepoint/client.py

# 2. Verify JSON-RPC stdio protocol & tool registration
PYTHONPATH=src python3 -c "
import subprocess, json
cmd = ['/home/sonnh95/.local/bin/mcp-auto-365-ms']
p = subprocess.Popen(cmd, stdin=subprocess.PIPE, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

# Handshake
p.stdin.write(json.dumps({'jsonrpc': '2.0', 'id': 1, 'method': 'initialize', 'params': {'protocolVersion': '2024-11-05', 'capabilities': {}, 'clientInfo': {'name': 'test', 'version': '1'}}}) + '\n')
p.stdin.flush()
json.loads(p.stdout.readline())

# List tools
p.stdin.write(json.dumps({'jsonrpc': '2.0', 'id': 2, 'method': 'tools/list', 'params': {}}) + '\n')
p.stdin.flush()
tools = json.loads(p.stdout.readline())['result']['tools']
print(f'Verified {len(tools)} tools registered.')
p.terminate()
"
```

---

## 7. Deployment & Multi-Harness Configuration

Run `./install.sh` after any launcher or dependency update. It automatically synchronizes configurations across:

1. **Claude Code (`~/.claude.json`):**
   ```json
   {
     "mcpServers": {
       "auto-365-ms": {
         "command": "/home/sonnh95/.local/bin/mcp-auto-365-ms"
       }
     }
   }
   ```
2. **Zed Editor (`~/.config/zed/settings.json`):**
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
3. **Oh My Pi (`~/.omp/agent/mcp.json`):**
   ```json
   {
     "mcpServers": {
       "auto-365-ms": {
         "command": "/home/sonnh95/.local/bin/mcp-auto-365-ms"
       }
     }
   }
   ```

---

## 8. Known Pitfalls & Troubleshooting Matrix

| Issue / Symptom | Root Cause | Resolution |
| :--- | :--- | :--- |
| **`HTTP Error 403: Forbidden` on Teams Graph API** | Microsoft Graph Protected APIs require tenant admin consent. | Do not use Graph for Teams. Use `TeamsAuthManager` and `apac.ng.msg.teams.microsoft.com` with `skypetoken_asm`. |
| **`HTTP Error 401: Unauthorized` on SharePoint Graph** | Azure CLI token expired or CAE challenged. | `SharePointClient.call_graph` catches 401 and automatically refreshes token via `az account get-access-token`. |
| **Direct URL download fails with 403** | Missing browser User-Agent or inverted cookie order. | Ensure headers have `rtFa` before `FedAuth` and User-Agent is set to `Mozilla/5.0 ... Chrome/153.0.0.0`. |
| **Tools list shows old schema after editing code** | Daemon process cached in RAM by harness. | Run `pkill -f "python3.*mcp-auto-365-ms/src/server.py"` to trigger a clean restart on next invocation. |
| **`Invalid format of emotions property`** | Teams Chat Service restricts reaction property updates on self-chat (`48:notes`). | Do not call reaction PUT on self-chat notes. Use standard messaging lifecycle. |
| **URL contains control characters error** | Spaces in SharePoint relative URLs not URL-encoded. | Use `urllib.parse.quote(path, safe='/:')` on all paths before constructing HTTP request URLs. |
