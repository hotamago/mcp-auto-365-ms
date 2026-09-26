# 🤖 AGENTS.md — Developer & AI Agent Operations Manual

> **Audience:** AI coding agents (Claude Code, Oh My Pi, Zed, Cursor) and engineers maintaining `mcp-auto-365-ms`.
> **Mission:** Autonomous access to Microsoft 365 (SharePoint, OneDrive, Teams, Outlook mail) from a coding harness, without tenant-admin Graph permissions and without degrading documents.

---

## 1. Design principles

1. **Never trust a single auth channel.** Graph and the browser session fail independently and for different reasons. Tools pick the channel that actually works for the job, and `check_365_connection` reports both.
2. **Fail loudly and usefully.** Every error carries a remediation line. A partially completed parallel scan must say so — a short answer that looks complete is worse than an error.
3. **Derive identity, never hardcode it.** User MRI, display name and aliases come from the session token.
4. **One declaration per tool.** Tools live in `src/tools.py` and are mounted onto each server.
5. **Do not summarise inside the server.** There is no model here. Gather structured data and expose prompts.

---

## 2. Layout

```text
mcp-auto-365-ms/
├── pyproject.toml            # uv project: deps, ruff, pytest
├── uv.lock                   # pinned, committed
├── .python-version           # 3.12
├── config.example.toml
├── install.sh                # uv-based installer
├── bin/                      # launchers -> `uv run python src/<server>.py`
├── src/
│   ├── server.py             # unified server (all 30 tools)
│   ├── tools.py              # single source of truth for tools/prompts/resources
│   ├── common/
│   │   ├── config.py         # env > user toml > repo toml > defaults
│   │   ├── errors.py         # typed errors + HTTP classification
│   │   ├── http.py           # timeouts, retry/backoff, one code path
│   │   ├── chrome_cookies.py # libsecret + cookie DB reader
│   │   ├── identity.py       # signed-in user, mention matching
│   │   └── health.py         # check_365_connection
│   ├── sharepoint/{client,server}.py
│   ├── teams/{auth,client,server}.py
│   └── outlook/{auth,client}.py
└── tests/                    # 371 offline tests
```

---

## 3. Authentication invariants

### 3.1 Browser cookies (Linux / libsecret)
- Profile is discovered from config (`browser.name`, `browser.profile`, `"auto"` supported).
- **The keyring item is located by attributes** (`application=chrome`), never by object path. The old `/login/2` path was an item *index* — installing any other Chromium app (Cursor, Termius, …) shifts it and yields the wrong key.
- `v10` values are encrypted with the fixed password `peanuts`; **only `v11` uses the keyring secret**.
- A 32-byte `sha256(host_key)` prefix is stripped only after verifying it matches.
- The DB is copied with its `-wal`/`-shm` sidecars into a private 0700 temp dir that is always removed.

### 3.2 Microsoft Teams
- **Chat Service:** `skypetoken_asm` → `https://{region}.ng.msg.teams.microsoft.com/v1`, header `Authentication: skypetoken=<jwt>`. Region comes from the `rgn` claim.
- **Middle tier:** `authtoken` cookie (stored as `Bearer=<jwt>&Origin=...`), audience `api.spaces.skype.com`. Used for calendar.
- **Identity:** the skypetoken's `skypeid` claim is `orgid:<guid>` — messages use `8:orgid:<guid>`. Always `normalize_mri()`.
- **Mentions:** the authoritative source is `properties.mentions`, a JSON array of `{itemid, mri, displayName, mentionType}`. ⚠️ The `itemid` in the HTML `<span>` is a **positional index into that array, not an MRI** — matching on it is always wrong.
- **Send returns no message id.** The response carries only `OriginalArrivalTime`; the id is recovered by matching `clientmessageid` in recent history.
- **Channels** end in `@thread.tacv2`; a thread reply targets `<channel-id>;messageid=<root>`.
- **Reactions:** `PUT/DELETE .../messages/{id}/properties?name=emotions`; `emotions` is a JSON-encoded `{key,value}` object inside the JSON body. Set `x-ms-client-caller` to `updateMessageReactionAdd`/`updateMessageReactionRemove`.
- Personal notes chat is `48:notes`.

### 3.3 SharePoint & OneDrive
- **Primary channel: Direct session (`rtFa=...; FedAuth=...`)** plus a browser User-Agent:
  - All operations (site/drive resolution, file downloads, search, version history, folder creation, file uploads and deletes) run natively against SharePoint's embedded `https://{host}/_api/v2.0/` and `/_api/web` endpoints.
  - State-changing requests (`POST`, `PUT`, `DELETE`) fetch and cache `FormDigestValue` from the **resource's own site** (`{site}/_api/contextinfo`) and are sent to that site. A digest from the host root is refused by `/sites/X` (403 on folder creation, 401 on upload).
  - Uploads ≤100 MB use REST v1 `{site}/_api/web/GetFolderByServerRelativeUrl('…')/Files/add(url='…',overwrite=true)`; `ensure_folder` GETs first and POSTs only the missing folders.
  - Immune to Azure CLI token expiration and Continuous Access Evaluation (CAE) disconnects.
- **Fallback channel: Graph (Azure CLI):** `az account get-access-token --resource https://graph.microsoft.com` is used as a secondary fallback if browser session cookies are unavailable.
- Always `urllib.parse.quote(path, safe='/:')` before building URLs.

### 3.4 Outlook mail
- Mail reuses the configured Chromium profile's persistent Microsoft sign-in cookies; no device login, separate Graph consent or app registration.
- `mail.client_id`, `tenant_id`, `login_host`, `origin`, `scope`, `redirect_uri`, `api_root` and optional `username` are config, not protocol constants hidden in client code. Defaults describe Outlook Web's public first-party deployment.
- Login cookies are replayed only to `mail.login_host`. Authorization is silent (`prompt=none`) with PKCE and state validation.
- The resulting audience must equal `mail.origin`. Keep only the short-lived access token in memory; discard the returned refresh token and mint again from browser cookies.
- Mail reads/writes use `mail.api_root`. If the browser session expires, tell the human to open `mail.origin` in the configured Chrome profile and choose **Stay signed in**.

---

## 4. Safety rules

<critical>
0. **NEVER SEND WITHOUT PER-MESSAGE HUMAN APPROVAL.** No Teams message, reaction, email, reply, edit, document comment or upload leaves this machine until the human has seen *that exact content and destination* and said yes. This is absolute and it outranks everything else in this file.
   - A standing instruction — "just send it", "do whatever you can", "go ahead" — is **not** approval of a draft that did not exist when it was said.
   - Approval for one message does **not** carry to the next one, not even in the same turn.
   - **In a group chat, a message meant for specific people MUST tag them** (`send_teams_message(mentions=[...])`). Busy groups bury untagged messages. The draft shown to the user must say who will be tagged.
   - A reaction is visible communication. Show the exact emoji/reaction, chat and message ID; require fresh approval before adding or removing it.
   - Risk rises: self-chat/self-email < 1:1 < group chat < company channel or external email. Group, channel and email sends need a fresh, explicit yes for that specific content and recipient set, every time.
   - The workflow is: compose → call with `is_user_confirm=false` to get the draft back → show the user the exact text and destination → wait for an explicit yes → call again with `is_user_confirm=true` (see §8).
   - **Why this is rule 0:** on 2026-09-21 an agent read "nhắn luôn đi" as blanket approval and posted five unreviewed questions into a squad channel containing the customer's BA and leads. It could not be taken back.
1. **TEST DESTINATIONS ONLY.** When testing Teams outbound tools, target only the personal self-chat (`48:notes`, `self`, `me`). When testing `send_email`, target only the signed-in user's own mailbox. Never target a colleague, group, channel or external address.
2. **CLEAN UP.** Delete every test message or email, and remove every test reaction, before ending your turn.
3. **`sync_folder_to_sharepoint` never deletes** and defaults to `dry_run=True`. Keep it that way.
4. **RESTART AFTER EDITS.** Editing Python does not reload a running daemon: `pkill -f "mcp-auto-365-ms/src/server.py"`.
5. **ONE SERVER REGISTRATION.** Only `auto-365-ms` in the harness configs. `doc-reader`/`teams-reader` expose subsets of the same tools and would duplicate them.
6. **NO NETWORK IN TESTS.** `tests/` must stay runnable offline.
</critical>

---

## 5. Verification

```bash
uv run ruff check src tests      # lint
uv run pytest -q                 # 371 offline tests

# Protocol smoke test: handshake + tool listing
uv run python - <<'PY'
import json, subprocess, time
p = subprocess.Popen(["./bin/mcp-auto-365-ms"], stdin=subprocess.PIPE,
                     stdout=subprocess.PIPE, text=True, bufsize=1)
def send(o): p.stdin.write(json.dumps(o) + "\n"); p.stdin.flush()
send({"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2024-11-05","capabilities":{},"clientInfo":{"name":"t","version":"1"}}})
p.stdout.readline()
send({"jsonrpc":"2.0","method":"notifications/initialized"})
send({"jsonrpc":"2.0","id":2,"method":"tools/list","params":{}})
while True:
    d = json.loads(p.stdout.readline())
    if d.get("id") == 2:
        print(f'{len(d["result"]["tools"])} tools'); break
p.terminate()
PY
```

Live behaviour is best checked with the `check_365_connection` tool.

---

## 6. Error handling contract

- Client code raises subclasses of `common.errors.Mcp365Error`, each with `message` + `remediation`.
- `tools.py` wraps every tool so those become `ToolError`. **This matters:** the MCP SDK passes a `ToolError` message through verbatim but replaces any other exception with a bare `"Error executing tool <name>"`, discarding the remediation.
- `common.http.request()` is the only outbound HTTP path: it enforces timeouts, retries `429/5xx` with jittered backoff, and classifies failures.
- Parallel scans return `(results, errors)`; tools render the errors via `_errors_note()`.
- Any other exception is also turned into a `ToolError` naming its type, message and `src/` file:line, and is logged with its traceback: the agent never gets a bare "Error executing tool".
- When the cookie channel and the Graph fallback both fail, the error lists **both** causes and both remediations; the Graph one (often a CAE 401) must never hide the real SharePoint one. `ConcurrentEditError` (409/412/423) never falls back to Graph.

---

## 7. Troubleshooting matrix

| Symptom | Root cause | Fix |
| :--- | :--- | :--- |
| `403` on `/_api` but web pages load | `FedAuth` is non-persistent (signed in without "Stay signed in"). Header `X-MSDAVEXT_Error: 917656`. | Sign in again with **Stay signed in** ticked. |
| Graph `401 TokenCreatedWithOutdatedPolicies` | Entra Continuous Access Evaluation challenge. | `az login --scope https://graph.microsoft.com/.default`. **Re-fetching the token does nothing** — the CLI returns the byte-identical cached token until real expiry. |
| Graph `403` on upload | Azure CLI token lacks `Files.*`/`Sites.*` scopes (tenant-dependent). | Check `check_365_connection`; re-login with the scope or ask an admin. |
| `KeyringError: không lấy được master key` | Keyring locked, or a different browser is configured. | Unlock the login keyring; set `MCP365_BROWSER`. |
| Teams tools 401 | skypetoken expired (~24h). | Reload `https://teams.microsoft.com` in Chrome. |
| Calendar tool 404s | The middle-tier calendar path is undocumented and version-dependent. | Override `teams.calendar_endpoint` in `config.toml`. |
| Outlook mail not connected | Chrome has no persistent Microsoft sign-in cookie, the selected account differs, or the session expired. | Open `mail.origin` (normally `https://outlook.office.com`) in the configured Chrome profile, select the intended account and choose **Stay signed in**. No device login or Graph consent is required. |
| Tools list shows stale schema | Daemon cached in RAM. | `pkill -f "mcp-auto-365-ms/src/server.py"`. |
| `ModuleNotFoundError: dbus` | Running with system Python instead of the uv env. | Use `uv run`, or the `bin/` launchers. |


---

## 8. User confirmation on every outbound tool

Every tool that sends, edits, deletes or overwrites takes a **required** `is_user_confirm`
argument (`common.approval.UserConfirm`). Its schema description — the same text on every tool —
tells the model it may only pass `true` after the user has seen that exact content and said yes.
Anything but a literal `true` refuses the call *before* any network request and returns the draft
(content + destination) so the model has something concrete to ask about.

| Tool | What the user approves |
| :--- | :--- |
| `send_teams_message`, `reply_to_channel_thread`, `edit_teams_message` | The exact text, the chat, and who will be tagged |
| `delete_teams_message` | Recalling that message |
| `react_to_teams_message` | The exact reaction, chat, message ID, and whether it is added or removed |
| `send_email` | Exact To/CC/BCC, subject and body |
| `upload_sharepoint_file` | The file and where it goes |
| `delete_sharepoint_item` | The exact path, size and item count, and Recycle Bin vs permanent |
| `sync_folder_to_sharepoint` | The upload plan (only when `dry_run=false`) |

That is the whole mechanism, on purpose: no destination is blocked and nothing is queued. Outbound
communication is sensitive, so the rule is one rule, everywhere — ask, then send.
`tests/test_approval.py` fails if a new outbound tool is added without the argument, or if a read-only tool grows one.

It is a contract with the calling model, not a cryptographic lock: a model that ignores the
description can still pass `true`. Rule 0 above is what the model is held to.

---

## 9. Conversation lookup

`fold()` in `src/teams/client.py` strips Vietnamese diacritics (NFD + an explicit `đ→d`,
which carries no combining mark) and casefolds. Both `list_conversations`' keyword filter and
all of `find_conversation`'s matchers run through it, so `nam son` finds
`1:1 Chat (Nguyễn Phan Nam Sơn)`.

⚠️ A 1:1 chat is named after the **other** member, found with `direct_chat_peer()` from the thread
id (`19:{guid1}_{guid2}@unq.gbl.spaces` names both members). The old code used `lastMessage.imdisplayname`,
so any 1:1 chat where the signed-in user spoke last was labelled with the user's own name — the chat
with Nguyễn Phan Nam Sơn read `1:1 Chat (Nguyễn Hoàng Sơn (VF-KPTX-VPTAITX))`. The 1:1 payload carries
no roster display names (`/threads/{id}` returns member MRIs only), so the name is learned from any
message that person sent anywhere on the conversations page, cached on the client per MRI and topped up
by every `read_messages`. Never resolved → the label is the peer's MRI, never the user's own name.
`last_sender` keeps reporting who actually spoke; only the label changed.

⚠️ `list_conversations` filters **before** truncating to `page_size`. The other order silently
hid every match outside the first page — a 1:1 chat 23 rows down was reported as
"không tìm thấy" for `filter_keyword="Nam Sơn", limit=15`. Do not reorder those two lines.


---

## 10. Resolving files anywhere in the tenant

Nothing may assume the configured site. `sharepoint.site_path` is a **fallback for bare GUIDs
only** — a GUID names no site, so it can only be looked up there.

| Concern | Rule |
| :--- | :--- |
| Site kinds | `_SITE_RE` matches `/sites/X`, `/teams/X` and `/personal/X`. A foreign host with no site segment is that host's root site — never the configured one. |
| Default library | `GET /sites/{id}/drive` (singular). `drives[0]` was just the first listed. |
| Library name | Never build `…/Shared Documents/…`. OneDrive's library is `Documents`; custom libraries have any name. Use `_item_file_url()`: Graph's `@microsoft.graph.downloadUrl`, else the drive's own `webUrl` + parent path. |
| Cookies | `FedAuth` is **per host**. `_cookie_headers(host=…)` must receive the host that serves the URL (`_download_headers()` does this). The old `%sharepoint.com%` lookup returned whichever host was used last. |
| Sharing links | `/:<letter>:/…` — Office letters on OneDrive go through WOPI; every other file link is fetched with `download=1`; `:f:` folders go to Graph. |
| Teams attachments | Not in the HTML body. `teams.client.parse_attachments()` reads `properties.files` (JSON string) → `objectUrl` on the sender's OneDrive. |

A missing `FedAuth` for a host means the browser never opened it. The fix is for the human to
open `https://<host>` in Chrome once — **never** to read cookie stores with an ad-hoc script.

## 11. Reading workbooks — no in-place edits

`read_sharepoint_sheet` downloads the file read-only and renders it with openpyxl
(`src/sharepoint/sheets.py`). **No tool edits an existing file in place.** `update_sharepoint_sheet`
(Graph workbook per-cell `PATCH`, whole-file fallback), `add_sharepoint_docx_comments` and
`replace_sharepoint_file` were removed on 26/09 at the user's request: writes into a file others had
open were refused (423 → 412 → 423 on `ViTa - S5 - Management Plan.xlsx`, 423 on Word files) or
had to re-upload the whole file. Do not bring them back; existing documents are edited in the
browser. `upload_sharepoint_file` and `sync_folder_to_sharepoint` stay — they add files (a
same-name file is replaced).

**Hidden sheets** (`hidden` and `veryHidden`) are skipped unless `include_hidden=True`: the listing
names them "(ẩn, bỏ qua)" and naming one raises a clear error. The author hid them on purpose; on
26/09 data from a hidden tab was wrongly taken as the source of truth.

**Merged cells.** A merged range keeps its value in the top-left (anchor) cell; every other cell of
the range is a `MergedCell` with a read-only `None` value and **no `column_letter`** — reading one
used to crash `render_sheet` on the first sheet with a merged banner. Column headers are built from
the column index, followers render blank, and the ranges are listed under the table.

Sheet names match exactly first, then ignoring case and surrounding spaces (`sheets.find_sheet`) —
real names carry trailing spaces (`'S5 Feature Release Plan '`).


## 12. Mentions

`send_teams_message(mentions=["Phạm Sỹ Hùng", ...])` tags people for real (a `<span itemtype=".../Mention">`
plus `properties.mentions`, JSON-encoded, `itemid` = position in that list).

`TeamsClient.resolve_mentions()` (shared by send and edit) takes, per entry: a name, `Name (Unit)`,
an MRI `8:orgid:<guid>`/bare GUID, an email/UPN, or an alias (`hoangnh21`). It matches the chat's own
history first (every message carries its sender's MRI, every mention the tagged person's MRI), then
the directory via `search_users()`:
- **MRI:** must appear in the chat history (the directory cannot look up an MRI; a silent member -> error asking for email).
- **Email/UPN, alias:** exact match on the directory entry's UPN/emails or their local part (`v.` prefix allowed).
  The People API misses a quoted full address, so the local part is what gets searched.
- **`Name (Unit)`:** exact (diacritic/case-insensitive) equality, unit included. Plain names: substring.
- **Namesakes** (different MRIs): narrowed to the chat's members (`get_members()`: `GET /threads/{id}`,
  MRIs only); still more than one -> error listing each candidate's name, email and MRI. Never guessed.
- **Display name** = the name the person most recently *sent* under. The 23/09 bug: a person who moved
  org unit keeps one MRI, and the old code kept the *first* (oldest) name, so the tag showed the old unit.
  Tag entries Teams split per word ("Nguyễn", "Hoàng") are dropped as aliases.
- Drafts show each tag as `@Name (email)` or `@Name (8:orgid:…last8)` (`mention_label()`).

Write `@Name` in the text to place the tag (the asked form, the display name, or it without `(Unit)`);
a person not written in the text is tagged at the start rather than silently dropped.

`edit_teams_message` builds tags with the same `render_message()`/`apply_mentions()` as send: an edit
replaces the whole message, so an `@Name` PUT as plain HTML is just text (the 23/09 bug). It takes the
same `mentions`; when omitted it reads the original and keeps each tag whose `@Name` (display name, or
without the `(Org unit)` suffix) is still in the new text (`kept_mentions()`), listing kept and dropped
tags in the draft. `mentions=[]` tags nobody. If the original cannot be read, nothing is edited.


## 13. Waiting for replies — `bin/mcp-365-watch`

An MCP server cannot push to the agent, so waiting is done by a **background CLI that exits
when something relevant arrives** — the exit is what wakes the agent (Claude Code:
`Bash(run_in_background=true)`).

```bash
bin/mcp-365-watch                                        # 1:1 messages + mentions of me
bin/mcp-365-watch --dm --mentions \
    --chat "[Vita-S5] Development team" --from "Phạm Sỹ Hùng" --timeout 1500
bin/mcp-365-watch --dm --mentions --interval 60 --min-interval 10 --half-life 300   # the defaults
```

- One conversation listing per poll; only chats with activity after the cursor are fetched, and in
  adaptive mode (default) a chat is not re-read while its `last_activity` is unchanged. The user's
  own messages never trigger it.
- **Adaptive pace.** Each chat has a heat = Σ weight × 0.5^(age / `--half-life`) over its recent
  messages: 1.0 for a 1:1 message (either direction) or a mention of me, 0.5 for my own message in a
  group or a `--from`-matching message in a `--chat`, 0 otherwise (a busy group that never tags me,
  `48:notes`). Own messages count — a reply is likely soon — but still never wake. Pace =
  `--interval / (1 + round(total heat))`, floored at `--min-interval`: 60 → 30 → 20 → 15 → 12 → 10 s
  by default, back to `--interval` once cold. One stderr line per pace change, never per poll.
- **No state file.** Heat is rebuilt from chat history: the first poll also reads up to `--warmup`
  (8) relevant chats active in the last 3 half-lives, so a watcher re-armed right after a wake-up
  starts fast. Those messages are older than `--since` and do not wake. `--no-adaptive` restores the
  fixed `--interval` pace with no warm-up and no re-read skipping.
- `--chat` + `--from` watch a busy group for specific people; `--dm` any 1:1; `--mentions` any tag.
  With no flags it defaults to `--dm --mentions`.
- Exit `0` = new messages printed · `3` = nothing within `--timeout` (re-arm) · `1` = auth/config
  error (printed to stdout so the agent is woken to tell the user).
- **Read-only.** Waking up is not permission to reply: every reply still goes through §0 / §8.

**Outlook: `bin/mcp-365-mail-watch`** — same contract for the Inbox, through
`OutlookMailClient.list_messages` (the `list_emails` code path, Chrome-session auth).

```bash
bin/mcp-365-mail-watch --from "nam son" --subject review --unread-only --timeout 1500
```

- Filters: `--from` (name or email), `--subject` (keyword), both diacritic-insensitive and
  repeatable; `--unread-only`, `--important`, `--flagged`. Different filters are ANDed.
- Output: VN time, sender, subject, one-line preview, message id (for `read_email`), and the exact
  re-arm arguments: `--since <newest shown − 2 min> --seen-ids '<ids shown in that window>'`
  (`watch_mail.rearm`). Re-arm with **both**. The old `newest + 1 s` lost a second mail in the
  same second and any mail delivered late with an older `ReceivedDateTime`; the overlapping window
  catches those and the id list keeps anything from printing twice. The cursor never goes back
  before the run's own `--since`. A timeout prints a re-arm line as well.
- After a poll with nothing new and a page that was not full, the cursor moves up to the newest
  mail examined − 2 min, so non-matching traffic cannot keep the window growing. A full page
  (50) leaves it alone: older mails in the window were never read.
- Exit codes as above. `AuthExpiredError`/`ConfigError`/`CookieError` exit `1` with the fix on
  stdout; every other `Mcp365Error` (network, timeout, 5xx after retries) skips one poll.
- Never marks mail read, moves, flags, deletes or sends. At most 50 newest mails per poll.


## 14. People search & inline image downloads

- **`find_user(query)`**: searches the company directory (Entra ID / Exchange GAL) via Outlook Web's
  People API (`/api/v2.0/me/people`). Supports full names, unaccented names, emails, aliases and phone
  numbers. Returns display name, email, UPN, job title, department, phone, Teams MRI (`8:orgid:<guid>`)
  and direct 1:1 chat ID (`19:{my_guid}_{their_guid}@unq.gbl.spaces`).
- **`download_message_images(chat_name_or_id, message_id=...)`**: downloads inline screenshots (AMSImage
  at `as-api.asm.skype.com` / `asyncgw.teams.microsoft.com` using `Cookie: skypetoken_asm=...`) and image
  attachments to local files, so coding agents can inspect them without custom Python scripts.
- **1:1 Direct Chat Provisioning & Canonical GUID Sorting**:
  - Teams Chat Service requires the two GUIDs in a 1:1 conversation thread (`19:{guid1}_{guid2}@unq.gbl.spaces`)
    to be lexicographically sorted (`sorted([my_guid, other_guid])`).
  - If two users have never chatted 1:1 before, Teams returns `404 LocationLookupFailed`. The client
    automatically provisions the conversation thread via `POST /v1/threads` (`create_or_get_direct_chat`).
  - `find_conversation` resolves 1:1 chats seamlessly from a reversed thread ID, user MRI (`8:orgid:<guid>`),
    or colleague name/email/UPN via directory search fallback.
