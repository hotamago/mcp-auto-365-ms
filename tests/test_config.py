"""Configuration precedence: env > user file > defaults."""

from __future__ import annotations

from common.config import get_config, reset_config_cache


def test_defaults_match_previous_hardcoded_values():
    """No config file must mean unchanged behaviour for the existing install."""
    cfg = get_config()
    assert cfg.sharepoint.hostname == "vingroupjsc.sharepoint.com"
    assert cfg.sharepoint.site_path == "/sites/VF_AIDV"
    assert cfg.sharepoint.site_url == "https://vingroupjsc.sharepoint.com/sites/VF_AIDV"
    assert cfg.sharepoint.site_name == "VF_AIDV"
    assert cfg.browser.name == "chrome"
    assert cfg.mail.client_id == "14d82eec-204b-4c2f-b7e8-296a70dab67e"
    assert cfg.mail.tenant_id == "organizations"


def test_env_overrides_defaults(monkeypatch):
    monkeypatch.setenv("MCP365_SHAREPOINT_HOSTNAME", "contoso.sharepoint.com")
    monkeypatch.setenv("MCP365_SHAREPOINT_SITE_PATH", "/sites/Engineering")
    reset_config_cache()
    cfg = get_config()
    assert cfg.sharepoint.site_url == "https://contoso.sharepoint.com/sites/Engineering"
    assert cfg.sharepoint.site_name == "Engineering"


def test_mail_env_overrides_defaults(monkeypatch):
    monkeypatch.setenv("MCP365_MAIL_CLIENT_ID", "custom-client")
    monkeypatch.setenv("MCP365_MAIL_TENANT_ID", "tenant-guid")
    reset_config_cache()
    cfg = get_config()
    assert cfg.mail.client_id == "custom-client"
    assert cfg.mail.tenant_id == "tenant-guid"


def test_numeric_env_is_cast(monkeypatch):
    monkeypatch.setenv("MCP365_HTTP_TIMEOUT", "12.5")
    monkeypatch.setenv("MCP365_MAX_WORKERS", "3")
    reset_config_cache()
    cfg = get_config()
    assert cfg.http.timeout == 12.5
    assert cfg.http.max_workers == 3


def test_malformed_numeric_env_falls_back(monkeypatch):
    monkeypatch.setenv("MCP365_HTTP_TIMEOUT", "not-a-number")
    reset_config_cache()
    assert get_config().http.timeout == 30.0


def test_mention_aliases_from_env(monkeypatch):
    monkeypatch.setenv("MCP365_MENTION_ALIASES", "sếp, boss ,")
    reset_config_cache()
    assert get_config().teams.extra_mention_aliases == ["sếp", "boss"]
