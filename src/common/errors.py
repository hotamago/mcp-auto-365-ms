"""Typed errors with actionable remediation messages.

Every failure mode this server hits in practice maps to a concrete user action.
Raising these (instead of returning ``"Error: ..."`` strings) lets the MCP layer
flag the call as failed *and* tells the operator exactly what to do next.
"""

from __future__ import annotations

import urllib.error
import urllib.parse


class Mcp365Error(Exception):
    """Base error. ``remediation`` is the concrete next step for the user.

    ``http_status`` / ``retry_after`` are filled in when the failure was an HTTP
    response. The Teams endpoint fallback needs the status itself: 503 and 429
    both map to :class:`RateLimitedError`, yet 503 means "try another endpoint"
    and 429 means "slow down, the endpoint is fine".
    """

    http_status: int | None = None
    retry_after: str | None = None

    def __init__(self, message: str, remediation: str = "") -> None:
        self.message = message
        self.remediation = remediation
        super().__init__(f"{message}\n→ {remediation}" if remediation else message)


class ConfigError(Mcp365Error):
    pass


class CookieError(Mcp365Error):
    """Chrome cookie store could not be read or decrypted."""


class KeyringError(CookieError):
    """GNOME Keyring / libsecret did not yield the browser master key."""


class AuthExpiredError(Mcp365Error):
    """A session token or cookie is present but no longer accepted."""


class CAEChallengeError(AuthExpiredError):
    """Entra Continuous Access Evaluation demands interactive re-auth.

    Critically: re-running ``az account get-access-token`` does NOT fix this.
    The Azure CLI returns the identical cached token until it actually expires,
    so a silent retry is a guaranteed no-op.
    """


class SharePointCookieRejectedError(AuthExpiredError):
    """FedAuth renders pages but is refused by ``/_api`` (error 917656).

    Happens when the SharePoint session was established without the
    "Stay signed in" / "Keep me signed in" option: the resulting cookie is
    non-persistent and SharePoint rejects it for REST and WebDAV calls.
    """


class RateLimitedError(Mcp365Error):
    pass


class NetworkError(Mcp365Error):
    """No HTTP response came back: the connection failed or broke off.

    The two subclasses matter for writes. ``ConnectError`` means the request
    never fully left this machine, so sending it again (to any host) cannot
    duplicate a message. ``TransportError`` means it was sent and the reply was
    lost, so the server may already have acted on it.
    """

    request_sent: bool = True


class ConnectError(NetworkError):
    """DNS, TCP connect, TLS handshake or send failed: nothing reached the server."""

    request_sent = False


class TransportError(NetworkError):
    """The request was sent but no complete response arrived (timeout, reset, EOF)."""

    request_sent = True


class ConversationNotFoundError(Mcp365Error):
    pass


class UnsupportedOperationError(Mcp365Error):
    """A capability that needs auth/scopes this deployment does not have."""


class HtmlPageError(Mcp365Error):
    """A download answered with a web page instead of the file's bytes.

    SharePoint does this, with HTTP 200, for ``Doc.aspx`` / sharing links it
    will not turn into a download (it serves the Office Online viewer) and when
    the session has lapsed (it serves the sign-in page). Writing that page under
    the file's name produced an ``.xlsx`` that was really HTML and looked like a
    successful download.
    """


class ConcurrentEditError(Mcp365Error):
    """The file changed (or is locked) since we read it; the write was refused.

    Raised on ``412 Precondition Failed``, or on ``409``/``423`` when a
    co-authoring session holds the file. The point is to fail instead of
    silently overwriting what a colleague just typed.
    """


_CAE_MARKERS = (
    "continuous access evaluation",
    "tokencreatedwithoutdatedpolicies",
    "interactionrequired",
)


def classify_http_error(exc: urllib.error.HTTPError, context: str = "") -> Mcp365Error:
    """Turn a raw ``HTTPError`` into a typed error with remediation.

    The response body is consumed here, so callers must not read it again.
    """
    try:
        body = exc.read().decode("utf-8", errors="ignore")
    except Exception:
        body = ""

    headers = getattr(exc, "headers", None)
    dav_error = ""
    if headers is not None:
        try:
            dav_error = urllib.parse.unquote_plus(headers.get("X-MSDAVEXT_Error", "") or "")
        except Exception:
            dav_error = ""

    where = f" khi {context}" if context else ""
    body_low = body.lower()

    # SharePoint: cookie valid for page rendering but refused for /_api.
    if "917656" in dav_error or "select the option to login automatically" in dav_error.lower():
        return SharePointCookieRejectedError(
            f"SharePoint refused the browser session cookie{where} (HTTP {exc.code}).\n"
            f"SharePoint says: {dav_error.strip()}",
            "Mở https://<tenant>.sharepoint.com trong Chrome, đăng xuất rồi đăng nhập lại và "
            "TICK 'Stay signed in' / 'Keep me signed in'. Cookie FedAuth không bền (non-persistent) "
            "vẫn hiển thị được trang web nhưng bị REST API từ chối.",
        )

    if any(marker in body_low for marker in _CAE_MARKERS):
        return CAEChallengeError(
            f"Microsoft Entra Continuous Access Evaluation đã thu hồi token{where} (HTTP {exc.code}).",
            "Chạy: az login --scope https://graph.microsoft.com/.default\n"
            "Lưu ý: gọi lại `az account get-access-token` KHÔNG giải quyết được — Azure CLI trả về "
            "đúng token cũ trong cache cho tới khi nó hết hạn thật.",
        )

    if exc.code == 401:
        return AuthExpiredError(
            f"Không được xác thực (HTTP 401){where}.",
            "Với SharePoint/Graph: chạy `az login`. Với Teams: mở lại https://teams.microsoft.com "
            "trong Chrome để làm mới skypetoken.",
        )

    if exc.code == 403:
        return AuthExpiredError(
            f"Bị từ chối truy cập (HTTP 403){where}." + (f"\nChi tiết: {dav_error.strip()}" if dav_error else ""),
            "Kiểm tra bạn có quyền trên site/tài nguyên này, và phiên đăng nhập Chrome còn hiệu lực. "
            "Chạy tool `check_365_connection` để xem kênh nào đang hỏng.",
        )

    if exc.code in (409, 412, 423):
        return ConcurrentEditError(
            f"File đã bị người khác sửa hoặc đang bị khoá{where} (HTTP {exc.code}). Chưa ghi gì cả.",
            "Không ghi đè. Chạy lại tool để tải bản mới nhất rồi áp lại thay đổi, "
            "hoặc đợi người đang mở file trên Excel/Word Online đóng lại.",
        )

    if exc.code == 429 or exc.code == 503:
        retry_after = ""
        if headers is not None:
            retry_after = headers.get("Retry-After", "") or ""
        return RateLimitedError(
            f"Bị giới hạn tần suất (HTTP {exc.code}){where}."
            + (f" Retry-After: {retry_after}s." if retry_after else ""),
            "Giảm `max_chats` / `limit`, hoặc thử lại sau ít phút.",
        )

    snippet = body.strip()[:400]
    return Mcp365Error(
        f"HTTP {exc.code} {exc.reason}{where}." + (f"\nPhản hồi: {snippet}" if snippet else ""),
        "Chạy `check_365_connection` để kiểm tra trạng thái các kênh xác thực.",
    )


class ApprovalRequiredError(Mcp365Error):
    """An outbound action was refused because the user has not approved it.

    Raised *before* anything leaves the machine. The message carries the draft;
    the agent is expected to show it to the user and wait for a yes.
    """
