"""Signed-in user identity, derived from the live session instead of hardcoded.

The original code matched mentions by comparing the rendered text against
hardcoded strings (``'@nguyễn hoàng sơn'``, ``'@sơn'``). That only worked for
one person and missed mentions whose display name was rendered differently.

Teams actually ships the authoritative data on every message:
``properties.mentions`` is a JSON array of
``{"itemid": 0, "mri": "8:orgid:<guid>", "displayName": ..., "mentionType": ...}``.
Note that the ``itemid`` in the HTML ``<span>`` is a positional *index* into
that array, not the MRI - so the span alone cannot identify the mentioned user.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field
from typing import Any

from .config import get_config


def normalize_mri(value: str) -> str:
    """Normalise a user id to full MRI form (``8:orgid:<guid>``).

    The skypetoken carries ``skypeid`` as ``orgid:<guid>`` (no ``8:`` prefix),
    while message payloads use the full ``8:orgid:<guid>``.
    """
    value = (value or "").strip()
    if not value:
        return ""
    if value.startswith("8:"):
        return value
    if value.startswith("orgid:"):
        return f"8:{value}"
    # A bare GUID.
    if len(value) == 36 and value.count("-") == 4:
        return f"8:orgid:{value}"
    return value


@dataclass
class Identity:
    mri: str = ""
    guid: str = ""
    display_name: str = ""
    upn: str = ""
    tenant_id: str = ""
    aliases: list[str] = field(default_factory=list)

    @classmethod
    def from_claims(cls, claims: dict[str, Any]) -> Identity:
        mri = normalize_mri(claims.get("skypeid", "") or "")
        guid = mri.split(":")[-1] if mri else ""
        display_name = claims.get("name", "") or ""
        upn = claims.get("upn") or claims.get("unique_name") or claims.get("username") or ""

        aliases: list[str] = []
        for candidate in (display_name, upn.split("@")[0] if upn else ""):
            if candidate:
                aliases.append(candidate.lower())

        # "Nguyễn Hoàng Sơn (VF-KPTX-VPTAITX)" -> also match the bare name and
        # the last word, which is how colleagues usually tag someone.
        if display_name:
            bare = display_name.split("(")[0].strip()
            if bare:
                aliases.append(bare.lower())
                parts = bare.split()
                if len(parts) > 1:
                    aliases.append(parts[-1].lower())

        aliases.extend(a.lower() for a in get_config().teams.extra_mention_aliases)

        seen: set[str] = set()
        unique = [a for a in aliases if a and not (a in seen or seen.add(a))]
        return cls(
            mri=mri,
            guid=guid,
            display_name=display_name,
            upn=upn,
            tenant_id=claims.get("tid", "") or "",
            aliases=unique,
        )

    # ------------------------------------------------------------- mentions

    def parse_mentions(self, message: dict[str, Any]) -> list[dict[str, Any]]:
        """Return the structured mention list attached to a raw Teams message."""
        props = message.get("properties") or {}
        raw = props.get("mentions")
        if not raw:
            return []
        if isinstance(raw, str):
            try:
                raw = json.loads(raw)
            except json.JSONDecodeError:
                return []
        return raw if isinstance(raw, list) else []

    def is_mentioned(self, message: dict[str, Any], cleaned_text: str = "") -> tuple[bool, str]:
        """Is the signed-in user mentioned? Returns ``(matched, reason)``.

        Precedence: authoritative MRI match, then broadcast mentions, then a
        display-name fallback for clients that omit the structured payload.
        """
        for mention in self.parse_mentions(message):
            mri = normalize_mri(str(mention.get("mri", "")))
            if mri and self.mri and mri == self.mri:
                return True, "mri"
            mention_type = str(mention.get("mentionType", "")).lower()
            if mention_type in ("everyone", "all", "team", "channel", "tag"):
                return True, f"broadcast:{mention_type}"

        text = (cleaned_text or "").lower()
        if text:
            for term in get_config().teams.broadcast_aliases:
                if term.lower() in text:
                    return True, f"broadcast:{term}"
            for alias in self.aliases:
                if f"@{alias}" in text:
                    return True, f"name:{alias}"
        return False, ""
