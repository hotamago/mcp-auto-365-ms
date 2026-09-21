# MCP Auto 365 MS

A unified **Model Context Protocol (MCP)** server that lets AI coding agents (Claude Code, Zed, Oh My Pi, Cursor) work with **Microsoft 365 — SharePoint, OneDrive, Microsoft Teams and Outlook mail** directly from the editor.

**31 tools · 3 prompts · 3 resources · Python 3.12 · managed with [uv](https://docs.astral.sh/uv/)**

---

## 🌟 Highlights

- **Works around Graph's Teams restrictions.** Microsoft gates Teams messages behind Protected APIs (`Chat.Read`, `ChannelMessage.Read.All`), which reject Azure CLI and developer tokens. This server reads your existing browser session instead (`skypetoken_asm`, `authtoken`, `FedAuth`, `rtFa`) via libsecret.
- **Outlook from the existing browser session.** Read and send through Outlook Web's own first-party session, bootstrapped from Chrome sign-in cookies. No device login, app registration, Graph consent or refresh-token cache. Every email requires approval of its exact recipients, subject and body.
- **No document degradation.** `.docx`, `.xlsx`, `.pptx` and PDFs are transferred as raw binaries — tables, formulas and diagrams stay intact.
- **Mentions matched by identity, not by name.** Detection uses the mention payload Teams attaches to each message (your user MRI), so it is correct regardless of how your display name is rendered — and it works for any user without editing code.
- **Self-diagnosing.** `check_365_connection` probes every auth channel and prints the exact fix for whatever is broken, instead of a bare `HTTP 403`.
- **Honest failures.** Tools raise typed errors that reach the agent as `isError` plus a remediation line; a partially scanned feed says so rather than looking complete.

---

## 🚀 Install

Only [uv](https://docs.astral.sh/uv/) is required — it provisions Python and every dependency.

```bash
curl -LsSf https://astral.sh/uv/install.sh | sh   # if you don't have uv yet
git clone <this-repo> && cd mcp-auto-365-ms
./install.sh
```

`install.sh` syncs the environment, verifies the server starts, links the launchers into `~/.local/bin`, and registers `auto-365-ms` with Claude Code, Zed and Oh My Pi (backing up each config first).

Then restart your editor and run the **`check_365_connection`** tool.

For Outlook mail, open **`https://outlook.office.com`** once in the configured Chrome profile and
choose **Stay signed in**. The server then renews short-lived Outlook tokens from that browser
session in memory; it never asks for a separate Graph permission grant.

### Day-to-day commands

```bash
uv sync --extra dev        # install dev dependencies
uv run pytest              # run the test suite (no network required)
uv run ruff check src tests
uv run python src/server.py   # run the server directly
```

---

## ⚙️ Configuration

Everything that used to be hardcoded is now configurable. Defaults match the previous behaviour, so no config file is required.

Precedence: **environment variables** → `~/.config/mcp-auto-365-ms/config.toml` → `./config.toml` → built-in defaults.

```toml
[sharepoint]
hostname = "contoso.sharepoint.com"
site_path = "/sites/Engineering"

[mail]
tenant_id = "organizations"  # or pin your tenant GUID

[browser]
name = "chrome"        # chrome | chromium | brave | edge
profile = "Default"    # or "auto" to pick the most recent profile

[http]
timeout = 30.0
max_workers = 6
```

Common environment overrides: `MCP365_SHAREPOINT_HOSTNAME`, `MCP365_SHAREPOINT_SITE_PATH`, `MCP365_BROWSER`, `MCP365_BROWSER_PROFILE`, `MCP365_HTTP_TIMEOUT`, `MCP365_MENTION_ALIASES`, `MCP365_MAIL_CLIENT_ID`, `MCP365_MAIL_TENANT_ID`.

See `config.example.toml` for every option.

---

## 🛠️ Tool catalogue

### SharePoint & OneDrive (12)

Any site kind works — `/sites/…`, `/teams/…` and personal OneDrive `/personal/…` (where every
Teams chat attachment lives). Cookies are picked **per host**: OneDrive (`tenant-my`) and team
sites carry separate `FedAuth` cookies.


| Tool | Purpose |
| --- | --- |
| `search_sharepoint_files` | Full-text document search via the SharePoint REST API |
| `read_sharepoint_link` | Folder tree, or document metadata + version history |
| `download_sharepoint_link` | Download original binaries, or a whole folder recursively. Direct paths, sharing links of any kind (`:u:` zip/json included), GUIDs |
| `upload_sharepoint_file` | Upload a file, creating missing parent folders |
| `replace_sharepoint_file` | Replace in place, creating a new version and keeping the link |
| `compare_sharepoint_versions` | Diff two versions, or a local file against SharePoint — on the file's own site |
| `read_sharepoint_sheet` | List a workbook's sheets, or render one as a table with A1 addresses |
| `update_sharepoint_sheet` | Edit cells (and clone a template sheet); requires `is_user_confirm`; uploaded with `If-Match` so a concurrent edit is never overwritten |
| `add_sharepoint_docx_comments` | Add Word review comments anchored to phrases; live mode via Word Companion Add-in (zero conflict, no file replace) or offline `If-Match` upload; requires `is_user_confirm` |
| `get_word_companion_status` | Check Word Companion Add-in bridge status and list open Word documents connected for live commenting |
| `sync_folder_to_sharepoint` | Upload new/changed files. One-way, never deletes, `dry_run` by default |
| `download_meeting_recordings` | Fetch Teams meeting recordings stored in SharePoint |


### 📝 Direct Word Commenting (Word Companion Add-in)

To comment on Word documents while they are actively open in Word Online or Word Desktop without causing eTag conflicts or lock errors (`HTTP 423`), use the included **Word Companion Add-in**:
1. The local bridge starts automatically on `https://127.0.0.1:3650` (configurable in `config.toml` or `MCP365_WORD_PORT`).
2. Open `https://127.0.0.1:3650/` in Chrome once and accept the self-signed dev certificate, or download the cert from `https://127.0.0.1:3650/cert.crt`.
3. Download the sideload manifest from `https://127.0.0.1:3650/manifest.xml` or run `mcp-word-companion manifest`.
4. In Word Online: go to **Home** &gt; **Add-ins** &gt; **More Settings** &gt; **Upload My Add-in**, and select `mcp-auto-365-word-manifest.xml`.
5. The **Auto 365 Companion** taskpane opens and connects. `add_sharepoint_docx_comments` will then insert comments directly into the live Word document via Word JavaScript API without replacing the file!
### Microsoft Teams (13)

| Tool | Purpose |
| --- | --- |
| `list_teams_chats` | Group chats, 1:1 chats, meeting chats and channels. Accent-insensitive keyword (`nam son` → `Nam Sơn`) and a `chat_type` filter |
| `read_teams_chat` | Message history, with optional Markdown export |
| `get_recent_team_messages` | Parallel activity feed across active conversations |
| `get_my_mentions` | Mentions with surrounding context, matched by user MRI |
| `get_new_mentions_since` | Cursor-based polling for new mentions |
| `search_teams_chat_messages` | Multi-keyword search across chats and channels |
| `send_teams_message` | Send, quote-reply, attach a local file, and **@mention people** (resolved from the chat's history) — requires `is_user_confirm` |
| `reply_to_channel_thread` | Reply *inside* a channel thread rather than starting a new one |
| `edit_teams_message` | Edit one of your own messages |
| `delete_teams_message` | Delete/recall one of your own messages |
| `react_to_teams_message` | Add or remove 👍/❤️/😂/😮/😢/😡 on a message when acknowledgement is enough; requires approval of the exact reaction and target |
| `download_chat_attachments` | Download paperclip attachments (`properties.files`) and SharePoint links from a chat; filter by `file_name` |
| `get_calendar_today` | Meetings and join links, via the Teams middle tier |

### Outlook mail (3)

Mail reuses the configured Chrome profile's persistent Microsoft sign-in session and Outlook Web's
public first-party SPA flow. All deployment values are configurable under `[mail]`; no secret is
embedded, no device login is required, and no refresh token is written to disk.

| Tool | Purpose |
| --- | --- |
| `list_emails` | List/filter/search a mailbox folder and return message IDs |
| `read_email` | Read one message's full body |
| `send_email` | Send plain-text mail; requires approval of exact To/CC/BCC, subject and body |

### User confirmation

Every tool that sends, edits, deletes or overwrites has a **required** `is_user_confirm` argument
whose description tells the model to ask the user first. Without `true` the tool sends nothing and
returns the draft to show the user. See [AGENTS.md §8](AGENTS.md).

### Cross-cutting (3)

| Tool | Purpose |
| --- | --- |
| `check_365_connection` | Diagnose every auth channel and print the fix |
| `extract_action_items` | Gather mentions + request-shaped messages as triage material |
| `get_daily_briefing` | Mentions, discussions, calendar and document updates in one call |

### Prompts & resources

Prompts `summarize_chat_thread`, `draft_standup` and `triage_mentions` gather the data and hand the reasoning to your agent's model — the server does no summarising of its own.

Resources: `teams://chats`, `teams://mentions/recent`, `m365://health`.

---

## 🔐 How authentication works

| Channel | Credential | Used for |
| --- | --- | --- |
| Teams Chat Service | `skypetoken_asm` cookie | All chat/channel reads and writes |
| Teams middle tier | `authtoken` cookie | Calendar |
| SharePoint direct | `rtFa` + `FedAuth` cookies | Downloads, REST search, version history |
| Microsoft Graph | Azure CLI token | Uploads, replace, drive metadata |
| Outlook Web | Chrome Microsoft sign-in cookies → short-lived in-memory Outlook token | List, search, read and send mail in the signed-in mailbox |

Browser credentials are decrypted locally with the key from your desktop keyring (libsecret).
Outlook access tokens remain in process memory only. Nothing is sent anywhere except to Microsoft.

**Three failure modes worth knowing:**

1. *SharePoint returns 403 on `/_api` while web pages load fine* — your session was created without "Stay signed in", so the `FedAuth` cookie is non-persistent and REST rejects it. Sign in again with that box ticked.
2. *Graph returns 401 with `TokenCreatedWithOutdatedPolicies`* — a Continuous Access Evaluation challenge. Re-running `az account get-access-token` will **not** help (the CLI returns the same cached token); run `az login --scope https://graph.microsoft.com/.default`.
3. *Outlook says the browser session is unavailable* — open `https://outlook.office.com` in the configured Chrome profile, select the intended account and choose **Stay signed in**. No device login or Graph consent is needed.

`check_365_connection` detects all three and tells you which applies.

---

## 🧪 Tests

```bash
uv run pytest        # 124 tests, no network, no keyring, no browser
```

Coverage includes error classification against responses captured from Microsoft, mention matching, timezone handling, HTTP retry/backoff, cookie decryption (v10 vs v11), config precedence, and a regression guard against the conversation-listing N+1.
