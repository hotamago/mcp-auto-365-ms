"""Shared HTTP layer: enforced timeouts, bounded retries, typed errors.

Every outbound call in this project goes through here. Previously there were 17
``urlopen`` calls and none of them passed ``timeout=``, so a single hung request
could pin a thread-pool worker forever and the tool would never return.
"""

from __future__ import annotations

import http.client
import http.cookiejar
import json
import random
import time
import urllib.error
import urllib.parse
import urllib.request
from typing import Any

from .config import get_config
from .errors import Mcp365Error, RateLimitedError, classify_http_error

#: Status codes worth retrying: throttling plus transient server faults.
_RETRYABLE = {429, 500, 502, 503, 504}


class _RedirectFragmentCaptured(Exception):
    def __init__(self, value: str) -> None:
        self.value = value


class _FragmentRedirectHandler(urllib.request.HTTPRedirectHandler):
    def __init__(self, key: str, expected: dict[str, str] | None = None) -> None:
        self.key = key
        self.expected = expected or {}

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        fragment = urllib.parse.parse_qs(urllib.parse.urlparse(newurl).fragment)
        values = fragment.get(self.key)
        matches = all(fragment.get(name) == [value] for name, value in self.expected.items())
        if values and matches:
            raise _RedirectFragmentCaptured(values[0])
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def _sleep_for(attempt: int, retry_after: str | None) -> float:
    if retry_after:
        try:
            return min(float(retry_after), 30.0)
        except (TypeError, ValueError):
            pass
    cfg = get_config().http
    # Exponential backoff with jitter, so parallel workers do not resynchronise.
    return min(cfg.backoff_base * (2**attempt) + random.uniform(0, 0.4), 20.0)


def request(
    url: str,
    *,
    headers: dict[str, str] | None = None,
    method: str = "GET",
    data: bytes | None = None,
    timeout: float | None = None,
    context: str = "",
    max_retries: int | None = None,
) -> tuple[int, bytes, Any]:
    """Perform an HTTP request, returning ``(status, body, response_headers)``.

    Raises a typed :class:`~common.errors.Mcp365Error` on failure.
    """
    cfg = get_config().http
    timeout = cfg.timeout if timeout is None else timeout
    retries = cfg.max_retries if max_retries is None else max_retries
    hdrs = dict(headers or {})
    hdrs.setdefault("User-Agent", cfg.user_agent)

    last_error: Exception | None = None

    for attempt in range(retries + 1):
        req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.status, resp.read(), resp.headers
        except urllib.error.HTTPError as exc:
            if exc.code in _RETRYABLE and attempt < retries:
                retry_after = None
                try:
                    retry_after = exc.headers.get("Retry-After")
                except Exception:
                    pass
                time.sleep(_sleep_for(attempt, retry_after))
                last_error = exc
                continue
            raise classify_http_error(exc, context) from exc
        except TimeoutError as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(_sleep_for(attempt, None))
                continue
            raise Mcp365Error(
                f"Request timed out after {timeout:.0f}s ({context or url}).",
                "Mạng chậm hoặc dịch vụ Microsoft đang lỗi. Tăng MCP365_HTTP_TIMEOUT nếu cần.",
            ) from exc
        except urllib.error.URLError as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(_sleep_for(attempt, None))
                continue
            raise Mcp365Error(
                f"Không kết nối được tới {url} ({exc.reason}).",
                "Kiểm tra kết nối mạng / VPN của công ty.",
            ) from exc
        except (ConnectionError, http.client.HTTPException) as exc:
            # The server hung up mid-exchange (RemoteDisconnected, reset). urlopen
            # only wraps connect-time failures in URLError, so these escaped as raw
            # tracebacks. A GET is safe to repeat; a POST may already have landed
            # (a sent message), so it is reported instead of silently resent.
            last_error = exc
            if method in ("GET", "HEAD") and attempt < retries:
                time.sleep(_sleep_for(attempt, None))
                continue
            raise Mcp365Error(
                f"Máy chủ ngắt kết nối giữa chừng ({context or url}): {exc!r}.",
                "Lỗi mạng tạm thời. Với thao tác gửi/ghi, kiểm tra đã gửi được chưa rồi mới thử lại.",
            ) from exc

    raise RateLimitedError(
        f"Đã thử lại {retries} lần nhưng vẫn thất bại ({context or url}): {last_error}",
        "Thử lại sau ít phút.",
    )


def request_json(url: str, **kwargs: Any) -> dict:
    """Perform a request and parse the body as JSON (``{}`` when empty)."""
    _status, body, _headers = request(url, **kwargs)
    if not body:
        return {}
    try:
        return json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise Mcp365Error(
            f"Phản hồi không phải JSON hợp lệ từ {url}: {body[:200]!r}",
            "Thường là do bị chuyển hướng sang trang đăng nhập — phiên Chrome đã hết hạn.",
        ) from exc


def request_bytes(url: str, **kwargs: Any) -> bytes:
    _status, body, _headers = request(url, **kwargs)
    return body


def capture_cookie(url: str, *, headers: dict[str, str], name: str, host: str, timeout: float | None = None) -> str:
    """Follow ``url``'s redirect chain and return cookie ``name`` set for ``host``.

    ``request()`` follows redirects too, but drops every ``Set-Cookie`` issued
    along the way. SharePoint's sign-in hand-off is exactly such a chain: a
    request carrying only the tenant-wide ``rtFa`` is bounced through
    ``/_forms/default.aspx``, which answers with a host-scoped ``FedAuth`` and
    redirects back. Returns ``""`` when the cookie never appears.
    """
    cfg = get_config().http
    jar = http.cookiejar.CookieJar()
    opener = urllib.request.build_opener(urllib.request.HTTPCookieProcessor(jar))
    hdrs = dict(headers)
    hdrs.setdefault("User-Agent", cfg.user_agent)
    try:
        with opener.open(
            urllib.request.Request(url, headers=hdrs), timeout=cfg.timeout if timeout is None else timeout
        ):
            pass
    except (urllib.error.HTTPError, urllib.error.URLError, TimeoutError):
        # A 401/403 at the end of the chain is normal for a bare root URL; the
        # cookie may already have been issued on an earlier hop.
        pass
    host = host.lower()
    for cookie in jar:
        domain = cookie.domain.lstrip(".").lower()
        if cookie.name == name and (host == domain or host.endswith("." + domain)):
            return cookie.value or ""
    return ""


def capture_redirect_fragment(
    url: str,
    *,
    cookies: dict[str, str],
    cookie_domain: str,
    key: str,
    headers: dict[str, str] | None = None,
    timeout: float | None = None,
    expected: dict[str, str] | None = None,
    context: str = "",
) -> str:
    """Follow redirects with host-scoped cookies and capture a fragment value.

    OAuth SPA authorization returns ``#code=...`` on the final redirect. URL
    fragments never reach the destination server, so a normal opener follows
    the redirect and loses the code. This handler stops at that boundary while
    keeping login cookies scoped to ``cookie_domain``.
    """
    cfg = get_config().http
    domain = cookie_domain.lstrip(".").lower()
    jar = http.cookiejar.CookieJar()
    for name, value in cookies.items():
        jar.set_cookie(
            http.cookiejar.Cookie(
                version=0,
                name=name,
                value=value,
                port=None,
                port_specified=False,
                domain=f".{domain}",
                domain_specified=True,
                domain_initial_dot=True,
                path="/",
                path_specified=True,
                secure=True,
                expires=None,
                discard=True,
                comment=None,
                comment_url=None,
                rest={},
                rfc2109=False,
            )
        )
    opener = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(jar), _FragmentRedirectHandler(key, expected)
    )
    request_headers = dict(headers or {})
    request_headers.setdefault("User-Agent", cfg.user_agent)
    try:
        with opener.open(
            urllib.request.Request(url, headers=request_headers),
            timeout=cfg.timeout if timeout is None else timeout,
        ):
            return ""
    except _RedirectFragmentCaptured as captured:
        return captured.value
    except urllib.error.HTTPError as exc:
        raise classify_http_error(exc, context) from exc
    except (urllib.error.URLError, TimeoutError) as exc:
        raise Mcp365Error(
            f"Không hoàn tất được redirect đăng nhập ({context or url}): {exc}",
            "Kiểm tra kết nối mạng / VPN của công ty.",
        ) from exc
