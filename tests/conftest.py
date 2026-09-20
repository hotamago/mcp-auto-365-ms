"""Shared fixtures. Tests never touch the network, the keyring or Chrome."""

from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

SRC = Path(__file__).resolve().parent.parent / "src"
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

from common.config import reset_config_cache  # noqa: E402


@pytest.fixture(autouse=True)
def _clean_config(monkeypatch):
    """Every test starts from pristine, env-free configuration."""
    for var in list(os.environ):
        if var.startswith("MCP365_"):
            monkeypatch.delenv(var, raising=False)
    reset_config_cache()
    yield
    reset_config_cache()


@pytest.fixture
def identity():
    from common.identity import Identity

    return Identity.from_claims(
        {
            "skypeid": "orgid:b6cf511d-9f31-4a84-89d8-3a400a1a544f",
            "name": "Nguyễn Hoàng Sơn (VF-KPTX-VPTAITX)",
            "upn": "sonnh95@vingroup.net",
            "tid": "ed6a2939-d153-4f92-94f8-3d790d96c9f8",
        }
    )
