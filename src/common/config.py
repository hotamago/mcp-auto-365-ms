"""Central configuration for mcp-auto-365-ms.

Resolution order (highest priority first):
  1. Environment variables (``MCP365_*``)
  2. ``~/.config/mcp-auto-365-ms/config.toml``
  3. ``config.toml`` next to the repository root
  4. Built-in defaults (identical to the values previously hardcoded in the clients)

Keeping the defaults equal to the old hardcoded constants means an existing
install keeps working with no config file present.
"""

from __future__ import annotations

import logging
import os
import tomllib
from dataclasses import dataclass, field
from functools import lru_cache
from pathlib import Path
from typing import Any

logger = logging.getLogger(__name__)

_REPO_ROOT = Path(__file__).resolve().parent.parent.parent
_USER_CONFIG = Path(os.environ.get("XDG_CONFIG_HOME", Path.home() / ".config")) / "mcp-auto-365-ms" / "config.toml"


def _load_toml(path: Path) -> dict[str, Any]:
    try:
        with path.open("rb") as fh:
            return tomllib.load(fh)
    except FileNotFoundError:
        return {}
    except Exception as exc:
        # A malformed config must never take the whole server down, but it must
        # not vanish silently either: every setting in it falls back to default.
        logger.warning("Ignoring unreadable config %s: %s", path, exc)
        return {}


@dataclass
class SharePointConfig:
    hostname: str = "vingroupjsc.sharepoint.com"
    site_path: str = "/sites/VF_AIDV"
    #: Folder used by ``send_teams_message(file_path=...)`` for chat attachments.
    attachment_folder: str = "VF-VSF Collaboration/00.ViTa/04. Squad Sharepoint/S5/02. Technical Docs/Shared from Chat"
    default_download_dir: str = "docs/sharepoint"

    @property
    def site_url(self) -> str:
        return f"https://{self.hostname}{self.site_path}"

    @property
    def site_name(self) -> str:
        return self.site_path.rstrip("/").split("/")[-1]


DEFAULT_CHAT_ENDPOINTS = (
    "https://teams.cloud.microsoft/api/chatsvc/{region}/v1",
    "https://teams.microsoft.com/api/chatsvc/{region}/v1",
    "https://{region}.ng.msg.teams.microsoft.com/v1",
)


@dataclass
class TeamsConfig:
    #: Extra names to treat as a mention of the user, on top of the ones derived
    #: from the signed-in identity. Matching is case-insensitive.
    extra_mention_aliases: list[str] = field(default_factory=list)
    #: Terms that count as a broadcast mention of everyone.
    broadcast_aliases: list[str] = field(default_factory=lambda: ["@all", "@everyone", "@team", "@channel"])
    #: Optional override for the Teams middle-tier region segment (normally
    #: derived from the skypetoken ``rgn`` claim).
    middle_tier_region: str = ""
    #: Middle-tier calendar endpoint template. Exposed as config because the
    #: path is undocumented and has changed between Teams releases.
    calendar_endpoint: str = "/api/mt/{region}/beta/me/calendarEvents?StartDate={start}&EndDate={end}"
    #: Chat Service base URLs, ``{region}`` = the skypetoken ``rgn`` claim. All
    #: three front the same service; the legacy ``*.ng.msg`` host started
    #: dropping TLS handshakes while the two web-app proxies stayed healthy, so
    #: every request can fail over between them (see ``teams/endpoints.py``).
    #: The list order is the preference used until latency has been measured.
    chat_endpoints: list[str] = field(default_factory=lambda: list(DEFAULT_CHAT_ENDPOINTS))
    #: Measure the endpoints (in parallel, lazily) and prefer the fastest.
    #: ``false`` keeps the configured order strictly (failover still applies).
    chat_probe: bool = True
    chat_probe_timeout: float = 4.0
    #: Seconds before the latency ranking is refreshed (in the background).
    chat_probe_ttl: float = 600.0
    #: Seconds an endpoint that just failed is demoted before it may lead again.
    chat_cooldown: float = 180.0
    #: Worst-case wall time for one logical Chat Service request, across all
    #: endpoints tried. Raised automatically when a tool passes ``timeout_seconds``.
    chat_budget: float = 25.0
    #: Attempts per request; they rotate through the ranked endpoints rather
    #: than hammering one host.
    chat_max_attempts: int = 3


@dataclass
class MailConfig:
    #: Outlook Web's first-party SPA settings. They are public identifiers,
    #: configurable because sovereign clouds and future Outlook deployments can
    #: use different hosts or clients.
    client_id: str = "9199bf20-a13f-4107-85dc-02114787ef48"
    tenant_id: str = "organizations"
    username: str = ""
    login_host: str = "login.microsoftonline.com"
    origin: str = "https://outlook.office.com"
    scope: str = "https://outlook.office.com/.default openid profile offline_access"
    redirect_uri: str = "https://outlook.office.com/mail/"
    api_root: str = "https://outlook.office.com/api/v2.0"


@dataclass
class BrowserConfig:
    #: Which Chromium-family browser to read cookies from.
    name: str = "chrome"
    #: Profile directory inside the browser's user-data dir. "auto" picks the
    #: most recently used profile that actually contains the needed cookies.
    profile: str = "Default"
    user_data_dir: str = ""


@dataclass
class HttpConfig:
    #: Default per-request timeout (seconds) for calls that are neither chat
    #: nor file transfer: Graph/SharePoint metadata, Outlook, sign-in.
    timeout: float = 30.0
    #: Per attempt, per endpoint, for the Teams Chat Service. A healthy
    #: endpoint answers in well under a second, so waiting longer than this
    #: mostly means the host is sick and the next endpoint is the better bet.
    timeout_chat: float = 10.0
    #: File downloads/uploads, recordings, attachments. The timeout applies to
    #: each socket read/write, not to the whole body, so large files still
    #: finish as long as bytes keep flowing.
    timeout_transfer: float = 120.0
    #: Hard ceiling for any timeout, including a tool's ``timeout_seconds``,
    #: so an agent can never make a request hang indefinitely.
    timeout_max: float = 600.0
    max_retries: int = 3
    backoff_base: float = 0.75
    max_workers: int = 6
    #: Conversation list cache TTL in seconds (kills the N+1 list_conversations storm).
    conversation_cache_ttl: float = 30.0
    user_agent: str = (
        "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36"
    )

    def timeout_for(self, kind: str) -> float:
        """Configured default timeout for a kind of work: chat, transfer or anything else."""
        return {"chat": self.timeout_chat, "transfer": self.timeout_transfer}.get(kind, self.timeout)


@dataclass
class Config:
    sharepoint: SharePointConfig = field(default_factory=SharePointConfig)
    teams: TeamsConfig = field(default_factory=TeamsConfig)
    mail: MailConfig = field(default_factory=MailConfig)
    browser: BrowserConfig = field(default_factory=BrowserConfig)
    http: HttpConfig = field(default_factory=HttpConfig)


def _to_bool(raw: str) -> bool:
    value = raw.strip().lower()
    if value in ("1", "true", "yes", "on"):
        return True
    if value in ("0", "false", "no", "off"):
        return False
    raise ValueError(raw)


def _split_list(raw: str) -> list[str]:
    return [item.strip() for item in raw.split(",") if item.strip()]


def _apply_section(obj: Any, values: dict[str, Any]) -> None:
    for key, value in values.items():
        if hasattr(obj, key):
            current = getattr(obj, key)
            # Keep declared types stable; ignore obviously wrong config types.
            if isinstance(current, list) and not isinstance(value, list):
                continue
            setattr(obj, key, value)


_ENV_MAP = {
    "MCP365_SHAREPOINT_HOSTNAME": ("sharepoint", "hostname", str),
    "MCP365_SHAREPOINT_SITE_PATH": ("sharepoint", "site_path", str),
    "MCP365_ATTACHMENT_FOLDER": ("sharepoint", "attachment_folder", str),
    "MCP365_DOWNLOAD_DIR": ("sharepoint", "default_download_dir", str),
    "MCP365_BROWSER": ("browser", "name", str),
    "MCP365_MAIL_CLIENT_ID": ("mail", "client_id", str),
    "MCP365_MAIL_TENANT_ID": ("mail", "tenant_id", str),
    "MCP365_MAIL_USERNAME": ("mail", "username", str),
    "MCP365_MAIL_LOGIN_HOST": ("mail", "login_host", str),
    "MCP365_MAIL_ORIGIN": ("mail", "origin", str),
    "MCP365_MAIL_SCOPE": ("mail", "scope", str),
    "MCP365_MAIL_REDIRECT_URI": ("mail", "redirect_uri", str),
    "MCP365_MAIL_API_ROOT": ("mail", "api_root", str),
    "MCP365_BROWSER_PROFILE": ("browser", "profile", str),
    "MCP365_BROWSER_USER_DATA_DIR": ("browser", "user_data_dir", str),
    "MCP365_HTTP_TIMEOUT": ("http", "timeout", float),
    "MCP365_HTTP_TIMEOUT_CHAT": ("http", "timeout_chat", float),
    "MCP365_HTTP_TIMEOUT_TRANSFER": ("http", "timeout_transfer", float),
    "MCP365_HTTP_TIMEOUT_MAX": ("http", "timeout_max", float),
    "MCP365_HTTP_MAX_RETRIES": ("http", "max_retries", int),
    "MCP365_MAX_WORKERS": ("http", "max_workers", int),
    "MCP365_CONVERSATION_CACHE_TTL": ("http", "conversation_cache_ttl", float),
    "MCP365_TEAMS_REGION": ("teams", "middle_tier_region", str),
    "MCP365_TEAMS_CHAT_ENDPOINTS": ("teams", "chat_endpoints", _split_list),
    "MCP365_TEAMS_CHAT_PROBE": ("teams", "chat_probe", _to_bool),
    "MCP365_TEAMS_CHAT_PROBE_TIMEOUT": ("teams", "chat_probe_timeout", float),
    "MCP365_TEAMS_CHAT_PROBE_TTL": ("teams", "chat_probe_ttl", float),
    "MCP365_TEAMS_CHAT_COOLDOWN": ("teams", "chat_cooldown", float),
    "MCP365_TEAMS_CHAT_BUDGET": ("teams", "chat_budget", float),
    "MCP365_TEAMS_CHAT_MAX_ATTEMPTS": ("teams", "chat_max_attempts", int),
}


@lru_cache(maxsize=1)
def get_config() -> Config:
    cfg = Config()

    for path in (_REPO_ROOT / "config.toml", _USER_CONFIG):
        data = _load_toml(path)
        for section in ("sharepoint", "teams", "mail", "browser", "http"):
            if isinstance(data.get(section), dict):
                _apply_section(getattr(cfg, section), data[section])

    for env_key, (section, attr, caster) in _ENV_MAP.items():
        raw = os.environ.get(env_key)
        if raw is None or raw == "":
            continue
        try:
            setattr(getattr(cfg, section), attr, caster(raw))
        except (TypeError, ValueError):
            continue

    aliases = os.environ.get("MCP365_MENTION_ALIASES", "")
    if aliases:
        cfg.teams.extra_mention_aliases = [a.strip() for a in aliases.split(",") if a.strip()]

    return cfg


def reset_config_cache() -> None:
    """Testing hook: drop the memoized config."""
    get_config.cache_clear()
