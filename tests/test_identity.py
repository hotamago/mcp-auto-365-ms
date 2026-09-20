"""Mention detection. Replaces the old hardcoded display-name matching."""

from __future__ import annotations

from common.identity import normalize_mri

OTHER_MRI = "8:orgid:de4cb212-40ef-4a20-81c4-246c0da5458e"


def test_normalize_mri_adds_missing_prefix():
    # The skypetoken carries "orgid:<guid>"; messages use "8:orgid:<guid>".
    assert normalize_mri("orgid:abc") == "8:orgid:abc"
    assert normalize_mri("8:orgid:abc") == "8:orgid:abc"
    assert normalize_mri("b6cf511d-9f31-4a84-89d8-3a400a1a544f") == "8:orgid:b6cf511d-9f31-4a84-89d8-3a400a1a544f"
    assert normalize_mri("") == ""


def test_identity_derives_aliases_from_claims(identity):
    # These were previously hardcoded strings in the client.
    assert identity.mri == "8:orgid:b6cf511d-9f31-4a84-89d8-3a400a1a544f"
    assert "sonnh95" in identity.aliases
    assert "sơn" in identity.aliases
    assert "nguyễn hoàng sơn" in identity.aliases


def test_mention_matched_by_mri(identity):
    message = {"properties": {"mentions": [{"itemid": 0, "mri": identity.mri, "mentionType": "person"}]}}
    matched, reason = identity.is_mentioned(message)
    assert matched and reason == "mri"


def test_mention_payload_as_json_string(identity):
    import json

    message = {"properties": {"mentions": json.dumps([{"itemid": 0, "mri": identity.mri, "mentionType": "person"}])}}
    assert identity.is_mentioned(message)[0]


def test_someone_elses_mention_is_not_mine(identity):
    message = {"properties": {"mentions": [{"itemid": 0, "mri": OTHER_MRI, "mentionType": "person"}]}}
    assert identity.is_mentioned(message, "@Trịnh Anh Tuấn xem giúp nhé")[0] is False


def test_itemid_index_is_not_treated_as_identity(identity):
    """`itemid` is a positional index into the mentions array, never an MRI."""
    message = {"properties": {"mentions": [{"itemid": 0, "mri": OTHER_MRI, "mentionType": "person"}]}}
    assert identity.is_mentioned(message)[0] is False


def test_broadcast_mention_matches(identity):
    matched, reason = identity.is_mentioned({}, "hi @all please review")
    assert matched and reason.startswith("broadcast")


def test_display_name_fallback_when_payload_missing(identity):
    matched, reason = identity.is_mentioned({}, "nhờ @Sơn check hộ")
    assert matched and reason.startswith("name:")


def test_plain_message_is_not_a_mention(identity):
    assert identity.is_mentioned({}, "deploy xong rồi nhé")[0] is False
