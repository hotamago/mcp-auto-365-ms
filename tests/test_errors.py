"""Error classification, built from responses actually captured from Microsoft."""

from __future__ import annotations

import email.message
import io
import urllib.error

from common.errors import (
    AuthExpiredError,
    CAEChallengeError,
    Mcp365Error,
    RateLimitedError,
    SharePointCookieRejectedError,
    classify_http_error,
)

#: Verbatim header returned by SharePoint for a non-persistent FedAuth session.
DAV_917656 = (
    "917656%3b+Access+denied.+Before+opening+files+in+this+location%2c+you+must+first+browse+"
    "to+the+web+site+and+select+the+option+to+login+automatically."
)

#: Verbatim body returned by Graph under a CAE challenge.
CAE_BODY = (
    b'{"error":{"code":"InvalidAuthenticationToken","message":"Continuous access evaluation resulted in '
    b'challenge with result: InteractionRequired and code: TokenCreatedWithOutdatedPolicies"}}'
)


def _http_error(code: int, body: bytes = b"{}", headers: dict[str, str] | None = None):
    msg = email.message.Message()
    for key, value in (headers or {}).items():
        msg[key] = value
    return urllib.error.HTTPError("https://example.invalid", code, "Err", msg, io.BytesIO(body))


def test_sharepoint_non_persistent_cookie_is_identified():
    err = classify_http_error(_http_error(403, headers={"X-MSDAVEXT_Error": DAV_917656}), "tìm kiếm")
    assert isinstance(err, SharePointCookieRejectedError)
    assert "Stay signed in" in err.remediation
    # The human-readable reason from SharePoint must survive URL-decoding.
    assert "login automatically" in err.message


def test_cae_challenge_is_identified_and_says_az_login():
    err = classify_http_error(_http_error(401, body=CAE_BODY))
    assert isinstance(err, CAEChallengeError)
    assert "az login" in err.remediation
    # Must warn that a silent token re-fetch cannot clear this.
    assert "KHÔNG" in err.remediation


def test_plain_401_is_auth_expired_but_not_cae():
    err = classify_http_error(_http_error(401, body=b'{"error":"expired"}'))
    assert isinstance(err, AuthExpiredError)
    assert not isinstance(err, CAEChallengeError)


def test_429_is_rate_limited_and_keeps_retry_after():
    err = classify_http_error(_http_error(429, headers={"Retry-After": "42"}))
    assert isinstance(err, RateLimitedError)
    assert "42" in err.message


def test_unknown_status_still_actionable():
    err = classify_http_error(_http_error(500, body=b"boom"))
    assert isinstance(err, Mcp365Error)
    assert err.remediation


def test_message_includes_context():
    err = classify_http_error(_http_error(403), context="tải tài liệu")
    assert "tải tài liệu" in err.message
