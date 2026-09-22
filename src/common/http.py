"""Shared HTTP layer: enforced timeouts, bounded retries, typed errors.

Every outbound call in this project goes through here. Previously there were 17
``urlopen`` calls and none of them passed ``timeout=``, so a single hung request
could pin a thread-pool worker forever and the tool would never return.

Timeouts depend on the kind of work. A Teams chat call that takes more than a
few seconds is a sick server, not a slow one, and is better failed over to
another endpoint; a recording download legitimately runs for minutes. Each
request resolves its timeout as: explicit ``timeout=`` argument (probes) →
the running tool's ``timeout_seconds`` (:func:`timeout_scope`) → the configured
default for its kind (``http.timeout_chat`` / ``timeout_transfer`` /
``timeout``), always capped by ``http.timeout_max``.
"""

from __future__ import annotations

import contextlib
import contextvars
import http.client
import http.cookiejar
import json
import random
import time
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Callable, Iterator
from pathlib import Path
from typing import Any

from .config import get_config
from .errors import ConnectError, HtmlPageError, Mcp365Error, RateLimitedError, TransportError, classify_http_error

#: Status codes worth retrying: throttling plus transient server faults.
_RETRYABLE = {429, 500, 502, 503, 504}

#: Methods that may be repeated after a timeout while reading the response.
#: POST is excluded: the server may already have acted (a sent message, an
#: e-mail), and a blind retry would do it twice.
_REPEATABLE_AFTER_TIMEOUT = {"GET", "HEAD", "PUT", "DELETE"}

_TIMEOUT_OVERRIDE: contextvars.ContextVar[float | None] = contextvars.ContextVar(
    "mcp365_timeout_override", default=None
)
_KIND: contextvars.ContextVar[str] = contextvars.ContextVar("mcp365_request_kind", default="default")


def _clamp(seconds: float) -> float:
    return max(1.0, min(float(seconds), get_config().http.timeout_max))


@contextlib.contextmanager
def timeout_scope(seconds: float | None) -> Iterator[None]:
    """Apply one command's ``timeout_seconds`` to every request made inside.

    Tools take the value from the agent; threading it through every client
    method signature would touch dozens of call sites, so it rides on a
    ``ContextVar`` instead. Worker pools must copy the context (see
    ``TeamsClient._scan``). ``None``/non-positive means "use the defaults".
    """
    if seconds is None or not isinstance(seconds, int | float) or seconds <= 0:
        yield
        return
    token = _TIMEOUT_OVERRIDE.set(_clamp(seconds))
    try:
        yield
    finally:
        _TIMEOUT_OVERRIDE.reset(token)


@contextlib.contextmanager
def kind_scope(kind: str) -> Iterator[None]:
    """Mark requests made inside as ``kind`` ("chat", "transfer") unless they say otherwise."""
    token = _KIND.set(kind)
    try:
        yield
    finally:
        _KIND.reset(token)


def timeout_override() -> float | None:
    """The running command's ``timeout_seconds`` (already capped), if any."""
    return _TIMEOUT_OVERRIDE.get()


def resolve_timeout(kind: str | None = None, explicit: float | None = None) -> float:
    """Timeout for one request: explicit → command override → configured default for ``kind``."""
    cfg = get_config().http
    if explicit is not None:
        return min(float(explicit), cfg.timeout_max)
    override = _TIMEOUT_OVERRIDE.get()
    if override is not None:
        return override
    return _clamp(cfg.timeout_for(kind or _KIND.get()))


def _timeout_hint(kind: str, method: str) -> str:
    """What to do about a timeout depends on what timed out."""
    if kind == "chat":
        # A tool-level ``timeout_seconds`` would not help here, so do not suggest it.
        return (
            "Teams Chat Service bình thường trả lời dưới 1 giây, nên đây gần như chắc chắn là máy chủ/mạng đang lỗi "
            "chứ không phải timeout quá ngắn. Thử lại sau ít phút; với thao tác gửi/ghi, kiểm tra đã gửi được chưa "
            "rồi mới thử lại."
        )
    written = (
        " Yêu cầu ghi có thể đã tới máy chủ: kiểm tra kết quả trước khi thử lại."
        if method not in ("GET", "HEAD")
        else ""
    )
    if kind == "transfer":
        cap = get_config().http.timeout_max
        return (
            "File lớn hoặc mạng chậm: gọi lại tool với `timeout_seconds` lớn hơn (ví dụ 300, tối đa "
            f"{cap:.0f}), hoặc đặt MCP365_HTTP_TIMEOUT_TRANSFER." + written
        )
    return (
        "Mạng chậm hoặc dịch vụ Microsoft đang lỗi. Nếu chắc là do mạng chậm, gọi lại tool với `timeout_seconds` "
        "lớn hơn (hoặc đặt MCP365_HTTP_TIMEOUT)." + written
    )


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


def _perform(
    url: str,
    *,
    headers: dict[str, str] | None,
    method: str,
    data: bytes | None,
    timeout: float | None,
    context: str,
    max_retries: int | None,
    kind: str | None,
    consume: Callable[[Any], Any],
) -> tuple[int, Any, Any]:
    """The retry loop shared by :func:`request` and :func:`request_to_file`.

    Error typing follows where urllib raised. ``do_open`` wraps only
    ``HTTPConnection.request()`` - DNS, TCP connect, TLS handshake and writing
    the request - in ``URLError``; ``getresponse()`` and body reads raise raw.
    So a ``URLError`` means the complete request never left this machine
    (:class:`ConnectError`, safe to resend anywhere), while a raw timeout, reset
    or TLS EOF means it was sent and only the reply was lost
    (:class:`TransportError`, a write may already have landed).
    """
    cfg = get_config().http
    kind = kind or _KIND.get()
    timeout = resolve_timeout(kind, timeout)
    retries = cfg.max_retries if max_retries is None else max_retries
    hdrs = dict(headers or {})
    hdrs.setdefault("User-Agent", cfg.user_agent)

    last_error: Exception | None = None

    for attempt in range(retries + 1):
        req = urllib.request.Request(url, data=data, headers=hdrs, method=method)
        try:
            with urllib.request.urlopen(req, timeout=timeout) as resp:
                return resp.status, consume(resp), resp.headers
        except urllib.error.HTTPError as exc:
            retry_after = None
            try:
                retry_after = exc.headers.get("Retry-After")
            except Exception:
                pass
            if exc.code in _RETRYABLE and attempt < retries:
                time.sleep(_sleep_for(attempt, retry_after))
                last_error = exc
                continue
            err = classify_http_error(exc, context)
            err.http_status = exc.code
            err.retry_after = retry_after
            raise err from exc
        except urllib.error.URLError as exc:
            last_error = exc
            if attempt < retries:
                time.sleep(_sleep_for(attempt, None))
                continue
            raise ConnectError(
                f"Không kết nối được tới {url} ({exc.reason}).",
                "Kiểm tra kết nối mạng / VPN của công ty.",
            ) from exc
        except TimeoutError as exc:
            last_error = exc
            if method in _REPEATABLE_AFTER_TIMEOUT and attempt < retries:
                time.sleep(_sleep_for(attempt, None))
                continue
            raise TransportError(
                f"Request timed out after {timeout:.0f}s ({context or url}).",
                _timeout_hint(kind, method),
            ) from exc
        except (ConnectionError, http.client.HTTPException, OSError) as exc:
            # The server hung up mid-exchange (RemoteDisconnected, reset, TLS
            # "UNEXPECTED_EOF_WHILE_READING"). urlopen only wraps send-side
            # failures in URLError, so these escaped as raw tracebacks. A GET is
            # safe to repeat; a POST may already have landed (a sent message), so
            # it is reported instead of silently resent.
            last_error = exc
            if method in ("GET", "HEAD") and attempt < retries:
                time.sleep(_sleep_for(attempt, None))
                continue
            raise TransportError(
                f"Máy chủ ngắt kết nối giữa chừng ({context or url}): {exc!r}.",
                "Lỗi mạng tạm thời. Với thao tác gửi/ghi, kiểm tra đã gửi được chưa rồi mới thử lại.",
            ) from exc

    raise RateLimitedError(
        f"Đã thử lại {retries} lần nhưng vẫn thất bại ({context or url}): {last_error}",
        "Thử lại sau ít phút.",
    )


def request(
    url: str,
    *,
    headers: dict[str, str] | None = None,
    method: str = "GET",
    data: bytes | None = None,
    timeout: float | None = None,
    context: str = "",
    max_retries: int | None = None,
    kind: str | None = None,
) -> tuple[int, bytes, Any]:
    """Perform an HTTP request, returning ``(status, body, response_headers)``.

    ``kind`` ("chat", "transfer", default) picks the default timeout and the
    advice given when it expires. Raises a typed
    :class:`~common.errors.Mcp365Error` on failure.
    """
    return _perform(
        url,
        headers=headers,
        method=method,
        data=data,
        timeout=timeout,
        context=context,
        max_retries=max_retries,
        kind=kind,
        consume=lambda resp: resp.read(),
    )


def decode_json(body: bytes, url: str) -> dict:
    """Parse a response body as JSON (``{}`` when empty), explaining login-page bodies."""
    if not body:
        return {}
    try:
        return json.loads(body.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise Mcp365Error(
            f"Phản hồi không phải JSON hợp lệ từ {url}: {body[:200]!r}",
            "Thường là do bị chuyển hướng sang trang đăng nhập — phiên Chrome đã hết hạn.",
        ) from exc


def request_json(url: str, **kwargs: Any) -> dict:
    """Perform a request and parse the body as JSON (``{}`` when empty)."""
    _status, body, _headers = request(url, **kwargs)
    return decode_json(body, url)


def request_bytes(url: str, **kwargs: Any) -> bytes:
    _status, body, _headers = request(url, **kwargs)
    return body


#: File types that legitimately *are* web pages; everything else is refused
#: when the server answers with HTML (see :func:`looks_like_html`).
HTML_SUFFIXES = (".html", ".htm", ".aspx", ".xhtml", ".mht", ".mhtml")


def looks_like_html(content_type: str, head: bytes) -> bool:
    """True when a response is a web page rather than a file's bytes.

    Checks the declared type *and* the first bytes: SharePoint sends its
    viewer and sign-in pages as ``text/html``, but a proxy or an older endpoint
    may label them ``application/octet-stream``.
    """
    if (content_type or "").split(";")[0].strip().lower() in ("text/html", "application/xhtml+xml"):
        return True
    start = head[:512].lstrip(b"\xef\xbb\xbf \t\r\n").lower()
    return start.startswith((b"<!doctype html", b"<html"))


def request_to_file(
    url: str,
    dest: Path | str,
    *,
    headers: dict[str, str] | None = None,
    context: str = "",
    timeout: float | None = None,
    max_retries: int | None = None,
    kind: str | None = "transfer",
    chunk_size: int = 1024 * 1024,
    reject_html: bool = False,
) -> int:
    """Stream a GET response body into ``dest``; returns the number of bytes written.

    Large files (meeting recordings run to gigabytes) are never held in memory.
    The body goes to ``<dest>.part`` and is renamed only once complete, so an
    interrupted download never leaves a truncated file under the real name. The
    timeout applies to each socket read, so a long download survives as long as
    bytes keep flowing.

    With ``reject_html`` an HTML answer raises :class:`HtmlPageError` before
    anything is written: the caller asked for a file, and a viewer or sign-in
    page saved under ``report.xlsx`` is worse than no file at all.
    """
    target = Path(dest)
    part = target.with_name(target.name + ".part")

    def consume(resp: Any) -> int:
        first = resp.read(chunk_size)  # network errors propagate to the retry loop
        if reject_html:
            content_type = (getattr(resp, "headers", None) or {}).get("Content-Type", "") or ""
            if looks_like_html(content_type, first):
                raise HtmlPageError(
                    f"Máy chủ trả về trang web (HTML) thay vì nội dung file '{target.name}' — đã không ghi file.",
                    "Link này trỏ tới trang xem/đăng nhập chứ không phải file. Dùng GUID (UniqueId) hoặc đường dẫn "
                    "đầy đủ tới file; nếu vẫn lỗi, mở link trong Chrome để làm mới phiên đăng nhập.",
                )
        written = 0
        try:
            fh = part.open("wb")  # truncates whatever an earlier attempt wrote
        except OSError as exc:
            raise Mcp365Error(f"Không ghi được file tạm {part}: {exc}", "Kiểm tra quyền ghi thư mục đích.") from exc
        with fh:
            chunk = first
            while chunk:
                try:
                    fh.write(chunk)
                except OSError as exc:
                    # Not a network fault: keep it out of the OSError retry branch.
                    raise Mcp365Error(
                        f"Không ghi được {target.name} ra đĩa: {exc}", "Kiểm tra dung lượng đĩa / quyền ghi."
                    ) from exc
                written += len(chunk)
                chunk = resp.read(chunk_size)  # network errors propagate to the retry loop
        return written

    try:
        _status, written, _headers = _perform(
            url,
            headers=headers,
            method="GET",
            data=None,
            timeout=timeout,
            context=context,
            max_retries=max_retries,
            kind=kind,
            consume=consume,
        )
    except BaseException:
        part.unlink(missing_ok=True)
        raise
    part.replace(target)
    return written


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
            urllib.request.Request(url, headers=hdrs), timeout=resolve_timeout(None, timeout)
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
            timeout=resolve_timeout(None, timeout),
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
