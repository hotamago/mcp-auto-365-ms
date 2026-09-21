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
│   ├── server.py             # unified server (all 29 tools)
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
└── tests/                    # 124 offline tests
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
- Personal notes chat is `48:notes`.

### 3.3 SharePoint & OneDrive
- **Graph (Azure CLI):** `az account get-access-token --resource https://graph.microsoft.com`. Used for writes and drive metadata.
- **Direct session:** `Cookie: rtFa=...; FedAuth=...` **plus a browser User-Agent** — required on *every* cookie call, including search and version history.
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
0. **NEVER SEND WITHOUT PER-MESSAGE HUMAN APPROVAL.** No Teams message, email, reply, edit, document comment or upload leaves this machine until the human has seen *that exact content and destination* and said yes. This is absolute and it outranks everything else in this file.
   - A standing instruction — "just send it", "do whatever you can", "go ahead" — is **not** approval of a draft that did not exist when it was said.
   - Approval for one message does **not** carry to the next one, not even in the same turn.
   - Risk rises: self-chat/self-email < 1:1 < group chat < company channel or external email. Group, channel and email sends need a fresh, explicit yes for that specific content and recipient set, every time.
   - The workflow is: compose → call with `is_user_confirm=false` to get the draft back → show the user the exact text and destination → wait for an explicit yes → call again with `is_user_confirm=true` (see §8).
   - **Why this is rule 0:** on 2026-09-21 an agent read "nhắn luôn đi" as blanket approval and posted five unreviewed questions into a squad channel containing the customer's BA and leads. It could not be taken back.
1. **TEST DESTINATIONS ONLY.** When testing Teams outbound tools, target only the personal self-chat (`48:notes`, `self`, `me`). When testing `send_email`, target only the signed-in user's own mailbox. Never target a colleague, group, channel or external address.
2. **CLEAN UP.** Delete every test message or email you send before ending your turn.
3. **`sync_folder_to_sharepoint` never deletes** and defaults to `dry_run=True`. Keep it that way.
4. **RESTART AFTER EDITS.** Editing Python does not reload a running daemon: `pkill -f "mcp-auto-365-ms/src/server.py"`.
5. **ONE SERVER REGISTRATION.** Only `auto-365-ms` in the harness configs. `doc-reader`/`teams-reader` expose subsets of the same tools and would duplicate them.
6. **NO NETWORK IN TESTS.** `tests/` must stay runnable offline.
</critical>

---

## 5. Verification

```bash
uv run ruff check src tests      # lint
uv run pytest -q                 # 124 offline tests

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

---

## 7. Troubleshooting matrix

| Symptom | Root cause | Fix |
| :--- | :--- | :--- |
| `403` on `/_api` but web pages load | `FedAuth` is non-persistent (signed in without "Stay signed in"). Header `X-MSDAVEXT_Error: 917656`. | Sign in again with **Stay signed in** ticked. |
| Graph `401 TokenCreatedWithOutdatedPolicies` | Entra Continuous Access Evaluation challenge. | `az login --scope https://graph.microsoft.com/.default`. **Re-fetching the token does nothing** — the CLI returns the byte-identical cached token until real expiry. |
| Graph `403` on upload/replace | Azure CLI token lacks `Files.*`/`Sites.*` scopes (tenant-dependent). | Check `check_365_connection`; re-login with the scope or ask an admin. |
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
| `send_teams_message`, `reply_to_channel_thread`, `edit_teams_message` | The exact text and the chat |
| `delete_teams_message` | Recalling that message |
| `send_email` | Exact To/CC/BCC, subject and body |
| `upload_sharepoint_file`, `replace_sharepoint_file` | The file and where it goes |
| `update_sharepoint_sheet` | The cell-by-cell change list |
| `add_sharepoint_docx_comments` | Each comment and the phrase it is anchored to |
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

## 11. Editing workbooks

`src/sharepoint/sheets.py` is pure (bytes in, bytes out). `update_sharepoint_sheet` reads the
file **and its eTag**, applies edits in memory, and stages the upload. On confirm it PUTs with
`If-Match: <eTag>`; Graph answers `412` if anyone saved since, and `409/423` while a
co-authoring session holds the file. All three map to `ConcurrentEditError` — the write is
refused, never forced. openpyxl drops charts and images on the round trip.
