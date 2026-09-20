# MCP Auto 365 MS (Model Context Protocol)

A unified, zero-configuration **Model Context Protocol (MCP)** server enabling AI coding agents (**Claude Code**, **Zed Editor**, **Oh My Pi - OMP**, **Cursor**) to seamlessly interact with **Microsoft 365 (SharePoint, OneDrive, and Microsoft Teams)** directly from their coding harness.

---

## 🌟 Key Highlights

- **Bypasses Microsoft Graph API Restrictions:** Microsoft enforces *Protected APIs* on Teams messages (`Chat.Read`, `ChannelMessage.Read.All`), blocking Azure CLI and standard developer tokens. `mcp-auto-365-ms` automatically retrieves and decrypts your active session tokens (`skypetoken_asm`, `FedAuth`, `rtFa`) directly from Google Chrome on Linux using GNOME Keyring (`org.freedesktop.secrets`).
- **One-Step Team Feed Aggregator:** No more multi-step friction (listing chats then opening each chat one by one). A single call fetches all fresh messages across all active group chats in parallel.
- **Smart Mention Detection with Context Window:** Extracts every discussion or task where colleagues or leads tag you (`@Nguyễn Hoàng Sơn`, `@Sơn`, `@all`), **kèm theo ngữ cảnh thảo luận xung quanh (các tin nhắn liền trước và liền sau)** để không bỏ lỡ yêu cầu chi tiết, tài liệu đính kèm, hay chỉ đạo liên quan.
- **Document Version Comparison (`compare_sharepoint_versions`):** So sánh trực tiếp 2 phiên bản tài liệu trên SharePoint hoặc giữa file local với SharePoint. Tự động trích xuất các đoạn thay đổi trong file Word (`.docx`), code/text/csv/markdown (`difflib.unified_diff`), và bảng tính (`.xlsx`).
- **Teams Channels & Threaded Discussions Support:** Tự động nhận diện và hiển thị tên chuẩn của các Team Channels (ví dụ: `[VF_VPTAITX] #Thông báo chung`) thay vì raw thread ID.
- **Generalized Message Dispatcher (`send_teams_message`):**
  - Gửi tin nhắn thông thường.
  - **Thread Quote-Reply:** Tự động tạo thẻ trích dẫn chuẩn `<blockquote ...>` trỏ thẳng vào tin nhắn cũ khi truyền `reply_to_id`.
  - **Tự động đính kèm file:** Tự động upload file local lên SharePoint và chèn liên kết file trực tiếp vào tin nhắn khi truyền `file_path`.
- **Chat Archiving & Markdown Export:** Hỗ trợ tham số `output_file` trong `read_teams_chat` để lưu toàn bộ biên bản trao đổi kỹ thuật ra file Markdown trên đĩa chỉ trong 1 lệnh.
- **Full Lifecycle Teams Messaging:** Hỗ trợ gửi, sửa (`edit_teams_message`), và xóa/thu hồi (`delete_teams_message`).
- **SharePoint Search & Versioned Management:**
  - `search_sharepoint_files`: Tìm kiếm tài liệu toàn diện trên SharePoint/OneDrive qua REST Search API (bỏ qua giới hạn Graph 403) chỉ trong ~1s.
  - `download_sharepoint_link`: Tải file/folder nhị phân gốc không bị suy giảm markdown.
  - `upload_sharepoint_file` & `replace_sharepoint_file`: Tải lên và cập nhật phiên bản mới (`Version 2.0`, `3.0`) trong khi giữ nguyên URL và Item ID.
- **Executive Daily Briefing (`get_daily_briefing`):** Tổng hợp tức thì đầu việc được tag, các cuộc thảo luận sôi nổi trong nhóm, và tài liệu cập nhật mới nhất cho buổi sáng làm việc.

---

## 🛠️ Complete Tool Catalog (16 Tools)

### 1. SharePoint & OneDrive Tools (6 Tools)
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `compare_sharepoint_versions` | **Smart Document Diff:** So sánh 2 phiên bản SharePoint hoặc so sánh file local vs SharePoint (Word `.docx`, text, code, `.xlsx`). | `file_a`: File local hoặc URL/GUID SharePoint<br/>`file_b`: (Tùy chọn) File thứ hai để đối chiếu<br/>`version_a`: Version cũ (nếu không truyền file_b)<br/>`version_b`: Version mới (mặc định: latest) |
| `search_sharepoint_files` | **Full Document Search:** Tìm kiếm tài liệu toàn diện trên SharePoint/OneDrive theo từ khóa hoặc định dạng file. Trả về UniqueId để download ngay. | `query`: Keyword tìm kiếm (e.g. `'SYS2'`, `'CAN'`, `'Architecture'`)<br/>`max_results`: Max kết quả (default: `20`)<br/>`file_extension`: Lọc đuôi file (e.g. `'docx'`, `'xlsx'`, `'pdf'`) |
| `read_sharepoint_link` | Explores a folder tree or inspects document metadata & version history. | `url`: SharePoint/OneDrive URL<br/>`max_depth`: Depth for folder traversal (default: `2`) |
| `download_sharepoint_link` | Downloads original binary files or entire folders preserving structure. | `url_or_guid`: File/Folder URL or unique GUID<br/>`target_dir`: Local directory (default: `docs/sharepoint`) |
| `upload_sharepoint_file` | **Uploader:** Uploads a local file to a SharePoint folder (auto-creates folder tree). | `local_file_path`: Path to local file<br/>`target_folder_url_or_path`: SharePoint URL or folder path<br/>`target_file_name`: Optional custom filename |
| `replace_sharepoint_file` | **Versioned Replacer:** Replaces an existing SharePoint file with a new version while preserving the item ID, URL, and recording version history. | `local_file_path`: Path to local updated file<br/>`file_url_or_guid`: File URL or GUID |

### 2. Microsoft Teams Automation Tools (10 Tools)
| Tool | Description | Key Parameters |
| :--- | :--- | :--- |
| `get_daily_briefing` | **Executive Morning Briefing:** Tự động tổng hợp 1-click toàn bộ nhiệm vụ tag bạn, trao đổi quan trọng trong team, và tài liệu mới cập nhật trong N giờ qua. | `hours`: Số giờ tổng hợp (default: `24`) |
| `get_my_mentions` | **Contextual Mention Tracker:** Trích xuất tin nhắn tag bạn, **kèm dải tin nhắn ngữ cảnh liền trước và liền sau** để nắm trọn vẹn yêu cầu công việc. | `hours`: Hours back (default: `72`)<br/>`limit`: Max mentions (default: `20`)<br/>`context_before`: Tin nhắn liền trước (default: `2`)<br/>`context_after`: Tin nhắn liền sau (default: `2`) |
| `get_recent_team_messages` | **One-Step Feed:** Scans all active group chats in parallel and returns all new messages within a time window. | `hours`: Hours back (default: `48`)<br/>`max_chats`: Max active chats to scan (default: `8`)<br/>`limit_per_chat`: Max messages per chat (default: `8`) |
| `send_teams_message` | **Generalized Message Dispatcher:** Gửi tin nhắn, trả lời trích dẫn (quote-reply), hoặc tự động đính kèm file local. Hỗ trợ `self`, `me`, `sonnh95@vingroup.net`. | `chat_name_or_id`: Chat name, email or thread ID<br/>`message`: Message text (markdown bold/code)<br/>`reply_to_id`: Optional ID tin nhắn cần trích dẫn reply<br/>`file_path`: Optional đường dẫn file local để upload & đính kèm |
| `read_teams_chat` | Đọc chi tiết hội thoại nhóm hoặc Channel. Hỗ trợ xuất biên bản lưu trữ ra file Markdown khi truyền `output_file`. | `chat_name_or_id`: Chat name/channel/thread ID<br/>`limit`: Number of messages (default: `30`)<br/>`since`: Date string (`today`, `yesterday`, `YYYY-MM-DD`)<br/>`only_mentions`: Filter to mentions only<br/>`output_file`: Tùy chọn đường dẫn file Markdown để lưu biên bản |
| `search_teams_chat_messages` | Tìm kiếm từ khóa kỹ thuật (hỗ trợ nhiều từ khóa cách nhau bởi dấu phẩy, e.g. `'CAN, S5, ZTNA'`) trên toàn bộ các nhóm chat & channels. | `query`: Từ khóa tìm kiếm đơn hoặc phân tách bằng dấu phẩy<br/>`limit`: Max results (default: `20`) |
| `list_teams_chats` | Lists recent group chats, 1:1 direct chats (including *Chat with yourself*), and meeting threads/channels (`[Team] #Channel`). | `limit`: Max chats (default: `30`)<br/>`filter_keyword`: Filter by name |
| `edit_teams_message` | **Message Editor:** Edits an existing message sent by you in a chat. | `chat_name_or_id`: Chat name or thread ID<br/>`message_id`: Message ID to edit<br/>`new_message`: Updated message text |
| `delete_teams_message` | **Message Deleter:** Deletes/revokes a message sent by you in a chat. | `chat_name_or_id`: Chat name or thread ID<br/>`message_id`: Message ID to delete |
| `download_chat_attachments` | **Teams-to-SharePoint Bridge:** Discovers SharePoint links in a chat and downloads the files directly. | `chat_name_or_id`: Chat name or thread ID<br/>`target_dir`: Target directory (default: `docs/sharepoint`) |

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

### Scenario 1: Review Specification Differences with Diff Tool
> **User Prompt:** *"So sánh xem file SYS2 mới sửa trên SharePoint khác gì so với bản nháp local của tôi."*  
> **Agent Action:** Calls `compare_sharepoint_versions(file_a="docs/SYS2_draft.docx", file_b="https://vingroupjsc.sharepoint.com/.../SYS2.docx")`  
> ➔ Trả về bảng tóm tắt: **+12 additions, -4 deletions** kèm diff rõ ràng từng đoạn văn bản.

### Scenario 2: Multi-Keyword Investigation & Export
> **User Prompt:** *"Tìm trong các nhóm xem có ai trao đổi về 'CAN, ZTNA, S5' không và xuất nội dung trao đổi quan trọng ra docs/can_notes.md."*  
> **Agent Action:**  
> 1. Calls `search_teams_chat_messages(query="CAN, ZTNA, S5")`.  
> 2. Calls `read_teams_chat(chat_name_or_id="Agent_in_8mem", output_file="docs/can_notes.md")`.

### Scenario 3: Executive Morning Briefing
> **User Prompt:** *"Tóm tắt nhanh tình hình công việc sáng nay giúp tôi."*  
> **Agent Action:** Calls `get_daily_briefing(hours=24)`  
> ➔ Báo cáo súc tích: Ai tag giao việc gì, nhóm nào đang bàn gì sôi nổi, tài liệu nào mới cập nhật.

---

## 🔒 Security & Privacy

- **Local Execution:** All decryptions occur purely in local memory on your machine.
- **Zero Cloud Leakage:** Tokens are never transmitted to any third-party service; requests communicate directly with official Microsoft endpoints (`https://*.sharepoint.com` and `https://*.teams.microsoft.com`).
- **Standard Linux Permissions:** Leverages standard GNOME Secret Service (`org.freedesktop.secrets`) matching your active user desktop session.

---

## 📄 License
MIT License.
