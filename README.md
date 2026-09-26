# MCP Auto 365 MS

A unified **Model Context Protocol (MCP)** server that lets AI coding agents (Claude Code, Zed, Oh My Pi, Cursor) work with **Microsoft 365 — SharePoint, OneDrive, Microsoft Teams and Outlook mail** directly from the editor.

**31 tools · 3 prompts · 3 resources · Python 3.12 · managed with [uv](https://docs.astral.sh/uv/)**

---

## 🌟 Highlights

- **Works around Graph's Teams restrictions.** Microsoft gates Teams messages behind Protected APIs (`Chat.Read`, `ChannelMessage.Read.All`), which reject Azure CLI and developer tokens. This server reads your existing browser session instead (`skypetoken_asm`, `authtoken`, `FedAuth`, `rtFa`) via libsecret.
- **Outlook from the existing browser session.** Read and send through Outlook Web's own first-party session, bootstrapped from Chrome sign-in cookies. No device login, app registration, Graph consent or refresh-token cache. Every email requires approval of its exact recipients, subject and body.
- **Cookie-first SharePoint & OneDrive writes.** File uploads, metadata, folder creation and version history run natively via browser cookies (`rtFa` + `FedAuth`) over SharePoint's embedded `_api/v2.0` and `_api/web`, eliminating frequent token expiry and CAE challenges from Azure CLI (which remains as a graceful fallback).
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
timeout = 30.0            # default per request
timeout_chat = 10.0       # Teams Chat Service, per attempt per endpoint
timeout_transfer = 120.0  # file downloads/uploads (streamed)
timeout_max = 600.0       # hard ceiling
max_workers = 6
```

Common environment overrides: `MCP365_SHAREPOINT_HOSTNAME`, `MCP365_SHAREPOINT_SITE_PATH`, `MCP365_BROWSER`, `MCP365_BROWSER_PROFILE`, `MCP365_HTTP_TIMEOUT`, `MCP365_HTTP_TIMEOUT_CHAT`, `MCP365_HTTP_TIMEOUT_TRANSFER`, `MCP365_HTTP_TIMEOUT_MAX`, `MCP365_TEAMS_CHAT_ENDPOINTS`, `MCP365_MENTION_ALIASES`, `MCP365_MAIL_CLIENT_ID`, `MCP365_MAIL_TENANT_ID`.

### Teams Chat Service endpoints

The Chat Service is reachable through several front doors: `teams.cloud.microsoft/api/chatsvc/{region}/v1`,
`teams.microsoft.com/api/chatsvc/{region}/v1` and the legacy `{region}.ng.msg.teams.microsoft.com/v1`. On first use
the server measures all of them in parallel and prefers the fastest (re-measured every 10 minutes in the background).
A connection failure, TLS error, timeout or 5xx moves the request to the next endpoint, and the failing one is demoted
for a cooldown. Sends, edits, reactions and deletions only switch endpoint when the request provably never reached the
server; if it may have landed, the tool says so instead of resending. `check_365_connection` shows the latency of each
endpoint and which one is in use; switches are logged to the MCP log (stderr). Order, probing and timings are
configurable under `[teams]` (`chat_endpoints`, `chat_probe`, `chat_cooldown`, `chat_budget`, ...).

### Per-call timeouts

Every networked tool accepts an optional `timeout_seconds`, applied to each HTTP request of that call and capped by
`http.timeout_max`. Chat tools default to 10 s per attempt: a slower chat call means a sick server, so the request fails
over rather than waiting. File tools default to 120 s per read/write and stream bodies to disk; pass a larger
`timeout_seconds` for big files, recordings or folders with many files.

See `config.example.toml` for every option.

---

## 🛠️ Tool catalogue

### SharePoint & OneDrive (10)

Any site kind works — `/sites/…`, `/teams/…` and personal OneDrive `/personal/…` (where every
Teams chat attachment lives). Cookies are picked **per host**: OneDrive (`tenant-my`) and team
sites carry separate `FedAuth` cookies.


| Tool | Purpose |
| --- | --- |
| `search_sharepoint_files` | Full-text document search via the SharePoint REST API |
| `read_sharepoint_link` | Folder tree, or document metadata + version history |
| `download_sharepoint_link` | Download original binaries, or a whole folder recursively. Direct paths, sharing links of any kind (`:u:` zip/json included), GUIDs |
| `upload_sharepoint_file` | Upload a file, creating missing parent folders; returns a link that opens it in the browser (`?web=1`) and the direct download link |
| `share_file_onedrive` | Upload a file to **your** OneDrive (`Shared from MCP`) and create a link anyone in the organization can open (view or edit) — for people who cannot open the team site; requires `is_user_confirm` |
| `compare_sharepoint_versions` | Diff two versions, or a local file against SharePoint — on the file's own site |
| `read_sharepoint_sheet` | List a workbook's sheets, or render one as a table with A1 addresses; hidden sheets are skipped unless `include_hidden` |
| `sync_folder_to_sharepoint` | Upload new/changed files. One-way, never deletes, `dry_run` by default |
| `delete_sharepoint_item` | Delete a file or folder (Recycle Bin by default); requires `is_user_confirm` |
| `download_meeting_recordings` | Fetch Teams meeting recordings stored in SharePoint |

No tool edits an existing file in place any more: cell edits, Word comments and whole-file replace
were removed on 26/09 — they were refused whenever someone had the file open (423/412) or had to
re-upload the whole file. Edit such files in the browser.

### Microsoft Teams (15)

| Tool | Purpose |
| --- | --- |
| `list_teams_chats` | Group chats, 1:1 chats, meeting chats and channels. Accent-insensitive keyword (`nam son` → `Nam Sơn`) and a `chat_type` filter |
| `read_teams_chat` | Message history, with optional Markdown export |
| `get_recent_team_messages` | Parallel activity feed across active conversations |
| `get_my_mentions` | Mentions with surrounding context, matched by user MRI |
| `get_new_mentions_since` | Cursor-based polling for new mentions |
| `search_teams_chat_messages` | Multi-keyword search across chats and channels |
| `send_teams_message` | Send, quote-reply, attach a local file, and **@mention people** by name, `Name (Unit)`, email/UPN, alias or MRI (chat history, then directory; namesakes refused with a candidate list) — requires `is_user_confirm` |
| `reply_to_channel_thread` | Reply *inside* a channel thread rather than starting a new one |
| `edit_teams_message` | Edit one of your own messages; keeps the original's @mentions still written as `@Name`, or takes `mentions` like send — requires `is_user_confirm` |
| `delete_teams_message` | Delete/recall one of your own messages |
| `react_to_teams_message` | Add or remove 👍/❤️/😂/😮/😢/😡 on a message when acknowledgement is enough; requires approval of the exact reaction and target |
| `download_chat_attachments` | Download paperclip attachments, inline images, and SharePoint links from a chat; filter by `file_name` |
| `download_message_images` | Download inline screenshots and image attachments from a chat or specific message ID directly to local files |
| `find_user` | Find a colleague in Microsoft 365 / Teams by name (with/without diacritics), email, alias, phone or keyword; returns contact info, Teams MRI and direct 1:1 chat ID |
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
| `get_daily_briefing` | Mentions, 1:1 messages, discussions, calendar and document updates in one call |

### Background watcher

`bin/mcp-365-watch` blocks until a new 1:1 message, a mention, or a message in a watched chat
arrives, prints it and exits — so an agent can wait for replies in the background instead of the
user prompting it. Read-only. See [AGENTS.md §13](AGENTS.md).

It polls every `--interval` seconds (60) while things are quiet and speeds up while a conversation
is hot — recent 1:1 messages, mentions of you and your own replies, decaying with `--half-life`
(300 s) — down to `--min-interval` (10 s), then drifts back. Heat is rebuilt from chat history, so a
watcher re-armed right after waking keeps the fast pace; `--no-adaptive` keeps the fixed pace.

Two levels, to wake the agent less often: 1:1 messages, mentions, quote replies to one of your
messages (`--replies-to-me`) and `--chat` messages wake it — after `--settle` more seconds to
collect the ones right behind (default 0 = at once). Everything else in a `--digest-chat` group is
collected and printed once the oldest has waited `--digest-after` seconds (60), or together with
the next wake-up. Output: `🔔 N tin mới`, then a `🔔 cần xem` and a `💬 tin nhóm` part.

`bin/mcp-365-mail-watch` does the same for Outlook: it polls the Inbox (default every 60 s) and
exits as soon as a mail received since `--since` matches the filters. Read-only: nothing is marked
as read, moved, flagged or deleted.

```bash
bin/mcp-365-mail-watch                                         # any new mail from now on
bin/mcp-365-mail-watch --from "nam son" --from alice@example.com --subject review --timeout 1500
bin/mcp-365-mail-watch --unread-only --important --since 2026-09-22T02:00:00Z
```

| Option | Meaning |
| --- | --- |
| `--since` | ISO time to watch from, UTC unless it has an offset (default: now) |
| `--seen-ids` | Comma-separated ids already shown; copy it from the re-arm line together with `--since` |
| `--from` | Sender name or email, diacritics optional (repeatable = any of them) |
| `--subject` | Keyword in the subject, diacritics optional (repeatable = any of them) |
| `--unread-only` / `--important` / `--flagged` | Only unread / High-importance / flagged mail |
| `--interval` / `--timeout` / `--scan` | Seconds between polls (60) / give up after (1500) / newest mails checked per poll (≤ 50) |

Different filters combine with AND. Each mail is printed with its Vietnam-time receipt, sender,
subject, a one-line preview and its id for `read_email`; the last line gives the exact arguments to
re-arm with: `--since` two minutes before the newest mail shown, plus `--seen-ids` for the mails
already shown in that window. A second mail in the same second, or one delivered late with an older
receive time, is still caught, and nothing is printed twice. A timeout (exit `3`) prints its own
re-arm line too; after a poll with nothing new the cursor moves up behind the newest mail examined.
Exit `0` = mail printed · `3` = nothing within `--timeout` · `1` = sign-in
or config problem (printed to stdout with the fix). Network errors, timeouts and 5xx skip that poll
and keep watching. It is a standalone CLI, so changing it needs no MCP server restart.

```text
📧 1 mail mới
- [22/09 11:56] Nam Sơn <nam.son@example.com> · Review kiến trúc S5 [📩 chưa đọc · 📎]
  > Anh xem giúp em bản kiến trúc mới, cần chốt trước thứ Sáu …
  id `AAMkADA5…`
→ Đọc: `read_email(message_id)`. Theo dõi tiếp: `--since 2026-09-22T04:54:30Z --seen-ids 'AAMkADA5…'`
```

### Prompts & resources

Prompts `summarize_chat_thread`, `draft_standup` and `triage_mentions` gather the data and hand the reasoning to your agent's model — the server does no summarising of its own.

Resources: `teams://chats`, `teams://mentions/recent`, `m365://health`.

---

## 🔐 How authentication works

| Channel | Credential | Used for |
| --- | --- | --- |
| Teams Chat Service | `skypetoken_asm` cookie | All chat/channel reads and writes |
| Teams middle tier | `authtoken` cookie | Calendar |
| SharePoint & OneDrive (Primary) | `rtFa` + `FedAuth` cookies | Reads, writes (uploads, folder creation, delete), downloads, search, version history |
| Microsoft Graph (Fallback) | Azure CLI token | Secondary fallback for drive operations when browser session is unavailable |
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
uv run pytest        # 398 tests, no network, no keyring, no browser
```

Coverage includes error classification against responses captured from Microsoft, mention matching, timezone handling, HTTP retry/backoff, cookie decryption (v10 vs v11), config precedence, and a regression guard against the conversation-listing N+1.
