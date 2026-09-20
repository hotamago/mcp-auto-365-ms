"""Shared HTTP layer: enforced timeouts, bounded retries, typed errors.

Every outbound call in this project goes through here. Previously there were 17
``urlopen`` calls and none of them passed ``timeout=``, so a single hung request
could pin a thread-pool worker forever and the tool would never return.
"""

from __future__ import annotations

import json
import random
import time
import urllib.error
import urllib.request
from typing import Any

from .config import get_config
from .errors import Mcp365Error, RateLimitedError, classify_http_error

#: Status codes worth retrying: throttling plus transient server faults.
_RETRYABLE = {429, 500, 502, 503, 504}


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
