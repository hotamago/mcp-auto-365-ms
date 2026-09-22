"""Teams Chat Service endpoints: use the fastest, fail over, cool down.

The same Chat Service answers on several front doors:

* ``https://teams.cloud.microsoft/api/chatsvc/{region}/v1`` (web-app proxy)
* ``https://teams.microsoft.com/api/chatsvc/{region}/v1`` (web-app proxy)
* ``https://{region}.ng.msg.teams.microsoft.com/v1`` (legacy direct host)

Only the last one used to be called. When it started dropping TLS handshakes
half of the time, every tool hung for 30 s x 4 attempts on that one host while
both proxies answered in under a second with the very same skypetoken.

So requests go through :class:`ChatServiceRouter`:

* **Fastest first.** On first use (lazily - never at import or construction)
  every endpoint is probed in parallel with a light authenticated GET; the
  first to answer is used right away and the stragglers refine the ranking in
  the background. The ranking is refreshed in the background after a TTL.
* **Fail over.** Network faults and 5xx move the request to the next endpoint.
  The failed endpoint is demoted for a cooldown (a short circuit breaker) and
  only leads again once it has cooled down and been re-measured.
* **Never duplicate a write.** POST/PUT/DELETE switch endpoints only when the
  request provably never left this machine (:class:`~common.errors.ConnectError`).
  If it was sent and the reply was lost, the caller is told to check before
  retrying - resending could post the message twice.
* **Business errors are not network errors.** 4xx (401 included) are returned
  as before; another front door would answer the same.
"""

from __future__ import annotations

import functools
import logging
import re
import threading
import time
import urllib.parse
from collections.abc import Callable
from dataclasses import dataclass
from typing import Any

from common.config import DEFAULT_CHAT_ENDPOINTS, get_config
from common.errors import ConnectError, Mcp365Error, NetworkError, TransportError
from common.http import resolve_timeout, timeout_override

logger = logging.getLogger(__name__)

#: Cheapest authenticated read that every front door serves (verified live).
PROBE_PATH = "/users/ME/conversations?view=msnp24Equivalent&pageSize=1"

#: Do not re-probe more often than this, even while an endpoint is recovering.
_MIN_REPROBE_INTERVAL = 30.0

Transport = Callable[..., tuple[int, bytes, Any]]

_MEDIA_HOST_SUFFIXES = ("asm.skype.com", "asyncgw.teams.microsoft.com", "ng.msg.teams.microsoft.com")


# ------------------------------------------------------------- URL helpers


def _host(base: str) -> str:
    return urllib.parse.urlsplit(base).netloc or base


def _short(err: BaseException) -> str:
    text = getattr(err, "message", "") or str(err)
    first = text.strip().splitlines()[0] if text.strip() else type(err).__name__
    return first if len(first) <= 160 else first[:157] + "..."


def _is_web_proxy(host: str) -> bool:
    host = host.lower()
    return host == "teams.microsoft.com" or host == "teams.cloud.microsoft" or host.endswith(".teams.cloud.microsoft")


def chat_templates() -> list[str]:
    """Configured endpoint templates, falling back to the defaults when unusable."""
    configured = [t.strip().rstrip("/") for t in get_config().teams.chat_endpoints if isinstance(t, str) and t.strip()]
    valid = [t for t in configured if t.startswith(("https://", "http://"))]
    if len(valid) != len(configured):
        logger.warning("Bỏ qua endpoint Chat Service không hợp lệ trong config: %s", sorted(set(configured) - set(valid)))
    return valid or list(DEFAULT_CHAT_ENDPOINTS)


def chat_bases(region: str) -> list[str]:
    """Concrete base URLs for ``region``, in configured order, without duplicates."""
    out: list[str] = []
    for template in chat_templates():
        base = template.replace("{region}", region)
        if base.lower() not in (b.lower() for b in out):
            out.append(base)
    return out


@functools.lru_cache(maxsize=64)
def _template_regex(template: str) -> re.Pattern[str]:
    escaped = re.escape(template.rstrip("/")).replace(re.escape("{region}"), r"(?P<region>[A-Za-z0-9-]+)")
    return re.compile(rf"^{escaped}(?P<rest>[/?#].*)?$", re.IGNORECASE)


def split_chat_url(url: str) -> tuple[str | None, str] | None:
    """``(region, path)`` for an absolute URL on any known Chat Service base, else ``None``.

    Links the service hands back (paging ``backwardLink``/``syncState`` URLs,
    ``Location`` headers, message URLs) may name any front door, typically the
    legacy host. Splitting them lets the request be re-routed through whichever
    endpoint is healthy now instead of the one that happened to be in the link.
    """
    templates = chat_templates()
    templates += [t for t in DEFAULT_CHAT_ENDPOINTS if t not in templates]
    for template in templates:
        match = _template_regex(template).match(url or "")
        if match:
            region = match.groupdict().get("region")
            rest = match.group("rest") or "/"
            return (region.lower() if region else None), rest
    return None


def is_teams_media_url(url: str) -> bool:
    """True for images/files served by Teams (AMS, async gateway, Chat Service or its proxies).

    Matching the host rather than a substring anywhere in the URL, and knowing
    the proxy front doors, keeps inline images working when message HTML starts
    pointing at ``teams.cloud.microsoft`` instead of the legacy hosts.
    """
    parts = urllib.parse.urlsplit(url or "")
    host = (parts.hostname or "").lower()
    if parts.scheme not in ("https", "http") or not host:
        return False
    if any(host == s or host.endswith("." + s) for s in _MEDIA_HOST_SUFFIXES):
        return True
    if split_chat_url(url):
        return True
    return _is_web_proxy(host) and parts.path.startswith("/api/")


def headers_for(base: str, headers: dict[str, str] | None) -> dict[str, str]:
    """Send the Origin a browser would send to this front door.

    The web-app proxies serve the Teams client itself, so the browser's calls
    to them are same-origin. Presenting ``teams.microsoft.com`` as the Origin
    to ``teams.cloud.microsoft`` would look like a cross-site request.
    """
    out = dict(headers or {})
    host = urllib.parse.urlsplit(base).hostname or ""
    if "Origin" in out and _is_web_proxy(host):
        origin = f"https://{host.lower()}"
        out["Origin"] = origin
        out["Referer"] = f"{origin}/"
    return out


# ---------------------------------------------------------------- routing


@dataclass
class _State:
    #: Round trip of the last successful probe, in seconds.
    latency: float | None = None
    #: Probe verdict: True answered, False answered with an unexpected HTTP
    #: status, None not measured since the last failure.
    reachable: bool | None = None
    cooldown_until: float = 0.0
    last_error: str = ""


class _ProbeRound:
    def __init__(self, bases: list[str], started_at: float, initial: bool) -> None:
        self.bases = bases
        self.started_at = started_at
        self.initial = initial
        self.outcomes: dict[str, str] = {}
        self.settled = threading.Event()  # first success, or everything answered
        self.done = threading.Event()


class ChatServiceRouter:
    """Process-wide endpoint ranking and failover for the Teams Chat Service.

    Shared by every ``TeamsClient`` in the process (the MCP server, the
    watcher, the health check): endpoint health is a fact about the network,
    not about one client object.
    """

    def __init__(
        self,
        *,
        clock: Callable[[], float] = time.monotonic,
        sleep: Callable[[float], None] = time.sleep,
        probe_enabled: bool | None = None,
    ) -> None:
        self._clock = clock
        self._sleep = sleep
        self._probe_enabled = probe_enabled
        self._lock = threading.Lock()
        self._states: dict[str, _State] = {}
        self._rounds: dict[tuple[str, ...], _ProbeRound] = {}
        self._primary: dict[tuple[str, ...], str] = {}

    # ------------------------------------------------------------- state

    def reset(self) -> None:
        with self._lock:
            self._states.clear()
            self._rounds.clear()
            self._primary.clear()

    def _state(self, base: str) -> _State:
        state = self._states.get(base)
        if state is None:
            state = self._states[base] = _State()
        return state

    def _probing(self) -> bool:
        return get_config().teams.chat_probe if self._probe_enabled is None else self._probe_enabled

    def _rank_locked(self, bases: list[str]) -> list[str]:
        now = self._clock()

        def key(item: tuple[int, str]) -> tuple[int, float, int]:
            index, base = item
            st = self._states.get(base) or _State()
            if st.cooldown_until > now:
                return (3, st.cooldown_until, index)  # last resort, soonest-recovered first
            if st.reachable is True and st.latency is not None:
                return (0, st.latency, index)
            if st.reachable is False:
                return (2, 0.0, index)
            return (1, 0.0, index)  # not measured yet: configured order

        return [base for _index, base in sorted(enumerate(bases), key=key)]

    def _note_primary_locked(self, bases: list[str], reason: str) -> None:
        key = tuple(bases)
        primary = self._rank_locked(bases)[0]
        previous = self._primary.get(key)
        if primary != previous:
            self._primary[key] = primary
            if previous:
                logger.warning("Teams Chat Service: đổi endpoint ưu tiên %s → %s (%s)", _host(previous), _host(primary), reason)
            else:
                logger.info("Teams Chat Service: dùng endpoint %s (%s)", _host(primary), reason)

    def rank(self, bases: list[str]) -> list[str]:
        with self._lock:
            return self._rank_locked(bases)

    def mark_failed(self, base: str, err: BaseException, bases: list[str] | None = None) -> None:
        cooldown = get_config().teams.chat_cooldown
        with self._lock:
            if bases:
                # Remember who led before this failure, so the switch is logged as one.
                self._primary.setdefault(tuple(bases), self._rank_locked(bases)[0])
            st = self._state(base)
            st.cooldown_until = self._clock() + cooldown
            st.reachable = None
            st.latency = None
            st.last_error = _short(err)
            if bases:
                self._note_primary_locked(bases, f"{_host(base)} lỗi: {st.last_error}")

    def mark_ok(self, base: str) -> None:
        with self._lock:
            st = self._state(base)
            st.cooldown_until = 0.0
            st.last_error = ""
            if st.reachable is not True:
                st.reachable = True

    def record_probe(self, base: str, latency: float | None, error: BaseException | None = None) -> str:
        """Store one probe outcome; returns a short human description of it."""
        cooldown = get_config().teams.chat_cooldown
        with self._lock:
            st = self._state(base)
            if error is None or getattr(error, "http_status", None) in (401, 403):
                # One sample is noisy (the same host measured 0.5 s and 4.4 s a
                # minute apart), so blend it with the previous measurement.
                smoothed = latency if st.latency is None or latency is None else 0.5 * st.latency + 0.5 * latency
            if error is None:
                st.latency, st.reachable, st.cooldown_until, st.last_error = smoothed, True, 0.0, ""
                return f"{latency:.2f}s"
            status = getattr(error, "http_status", None)
            if status in (401, 403):
                # The host answered; the session is the problem, and every
                # front door would say the same. Rank it by speed as usual.
                st.latency, st.reachable, st.last_error = smoothed, True, _short(error)
                return f"{latency:.2f}s (HTTP {status})"
            st.latency = None
            st.last_error = _short(error)
            if status is not None and status < 500:
                st.reachable = False  # answers, but not this API: rank after the healthy ones
                return f"HTTP {status}"
            st.reachable = None
            st.cooldown_until = self._clock() + cooldown
            return f"lỗi ({st.last_error})"

    # ------------------------------------------------------------ probing

    def _start_round_locked(
        self, bases: list[str], headers: dict[str, str], transport: Transport, *, initial: bool
    ) -> _ProbeRound:
        round_ = _ProbeRound(list(bases), self._clock(), initial)
        self._rounds[tuple(bases)] = round_
        pending = {"n": len(bases)}
        pending_lock = threading.Lock()
        timeout = get_config().teams.chat_probe_timeout

        def probe(base: str) -> None:
            started = self._clock()
            error: BaseException | None = None
            try:
                transport(
                    base + PROBE_PATH,
                    method="GET",
                    headers=headers_for(base, headers),
                    timeout=timeout,
                    max_retries=0,
                    context="đo độ trễ Teams Chat Service",
                    kind="chat",
                )
            except Exception as exc:  # noqa: BLE001 - any failure is a verdict on this endpoint
                error = exc
            outcome = self.record_probe(base, self._clock() - started, error)
            with pending_lock:
                round_.outcomes[base] = outcome
                pending["n"] -= 1
                finished = pending["n"] == 0
            if error is None or getattr(error, "http_status", None) in (401, 403):
                round_.settled.set()
            if finished:
                with self._lock:
                    self._note_primary_locked(round_.bases, "đo độ trễ")
                    ranking = self._rank_locked(round_.bases)
                logger.info(
                    "Teams Chat Service: đo độ trễ %s → thứ tự %s",
                    " · ".join(f"{_host(b)} {round_.outcomes.get(b, '?')}" for b in round_.bases),
                    " > ".join(_host(b) for b in ranking),
                )
                round_.settled.set()
                round_.done.set()

        # Daemon threads: a probe stuck in a TLS handshake must never keep the
        # process alive (the watcher exits right after printing), and they
        # deliberately do not inherit the tool's timeout override.
        for base in bases:
            threading.Thread(target=probe, args=(base,), name="chatsvc-probe", daemon=True).start()
        return round_

    def ordered(self, bases: list[str], headers: dict[str, str], transport: Transport) -> list[str]:
        """Endpoints best-first, probing lazily on first use and refreshing after the TTL."""
        if len(bases) < 2 or not self._probing():
            return self.rank(bases)
        cfg = get_config().teams
        key = tuple(bases)
        wait_for: _ProbeRound | None = None
        with self._lock:
            round_ = self._rounds.get(key)
            now = self._clock()
            if round_ is None:
                wait_for = self._start_round_locked(bases, headers, transport, initial=True)
            elif round_.initial and not round_.settled.is_set():
                wait_for = round_  # another thread started the first probe: share it
            elif round_.done.is_set():
                age = now - round_.started_at
                recovering = any(
                    (st := self._states.get(b)) is not None and st.reachable is None and st.cooldown_until <= now
                    for b in bases
                )
                if age >= cfg.chat_probe_ttl or (recovering and age >= _MIN_REPROBE_INTERVAL):
                    # Refresh in the background; this request uses the current ranking.
                    self._start_round_locked(bases, headers, transport, initial=False)
        if wait_for is not None:
            wait_for.settled.wait(cfg.chat_probe_timeout + 0.5)
        return self.rank(bases)

    def probe_now(self, bases: list[str], headers: dict[str, str], transport: Transport) -> list[dict[str, Any]]:
        """Measure every endpoint now and wait for all of them (health check)."""
        with self._lock:
            round_ = self._start_round_locked(bases, headers, transport, initial=False)
        round_.done.wait(get_config().teams.chat_probe_timeout + 1.0)
        return self.snapshot(bases)

    def snapshot(self, bases: list[str]) -> list[dict[str, Any]]:
        """Current view per endpoint, best first (for logs and ``check_365_connection``)."""
        with self._lock:
            now = self._clock()
            out = []
            for base in self._rank_locked(bases):
                st = self._states.get(base) or _State()
                out.append(
                    {
                        "base": base,
                        "host": _host(base),
                        "latency": st.latency,
                        "reachable": st.reachable,
                        "cooldown_left": max(0.0, st.cooldown_until - now),
                        "last_error": st.last_error,
                    }
                )
            return out

    # ------------------------------------------------------------ requests

    def call(
        self,
        region: str,
        path_or_url: str,
        *,
        method: str = "GET",
        headers: dict[str, str],
        data: bytes | None = None,
        context: str = "",
        transport: Transport,
    ) -> tuple[int, bytes, Any]:
        """Send one Chat Service request, failing over between endpoints.

        ``path_or_url`` is a path relative to the ``/v1`` base, or an absolute
        URL the service returned; one on a known front door is re-routed, any
        other absolute URL is fetched as-is.
        """
        method = method.upper()
        if "://" in path_or_url:
            split = split_chat_url(path_or_url)
            if split is None:
                return transport(path_or_url, method=method, headers=headers, data=data, context=context, kind="chat")
            url_region, path = split
        else:
            url_region, path = None, path_or_url
        if not path.startswith(("/", "?")):
            path = "/" + path
        bases = chat_bases(url_region or region)

        cfg = get_config()
        order = self.ordered(bases, headers, transport)
        per_attempt = resolve_timeout("chat")
        attempts = max(1, int(cfg.teams.chat_max_attempts))
        budget = cfg.teams.chat_budget
        if timeout_override() is not None:
            # The tool asked for a longer wait per attempt; give every attempt
            # that much instead of letting the default budget cut it short.
            budget = min(cfg.http.timeout_max, max(budget, per_attempt * attempts))
        deadline = self._clock() + budget
        replayable = method in ("GET", "HEAD")
        failures: list[tuple[str, Mcp365Error]] = []

        for n in range(attempts):
            base = order[n % len(order)]
            remaining = deadline - self._clock()
            if n and remaining < 1.0:
                break
            if n >= len(order):
                # Back on a host that already failed in this request: pause briefly.
                pause = min(0.5 * (2 ** (n - len(order))), remaining / 4)
                self._sleep(pause)
                remaining = deadline - self._clock()
            timeout = max(1.0, min(per_attempt, remaining))
            try:
                result = self._attempt(base, path, method, headers, data, context, transport, timeout, deadline)
            except Mcp365Error as err:
                if not _is_endpoint_fault(err):
                    raise
                self.mark_failed(base, err, bases)
                failures.append((base, err))
                was_sent = isinstance(err, TransportError) or err.http_status is not None
                if was_sent and not replayable:
                    raise self._outcome_unknown(base, err, context) from err
                nxt = order[(n + 1) % len(order)] if n + 1 < attempts else None
                logger.warning(
                    "Teams Chat Service: %s lỗi khi %s (%s)%s",
                    _host(base),
                    context or method,
                    _short(err),
                    f" → thử {_host(nxt)}" if nxt else "",
                )
                continue
            self.mark_ok(base)
            if failures:
                logger.warning(
                    "Teams Chat Service: %s thành công qua %s sau khi %s lỗi",
                    context or method,
                    _host(base),
                    ", ".join(_host(b) for b, _e in failures),
                )
            return result
        raise self._exhausted(failures, context, replayable)

    def _attempt(
        self,
        base: str,
        path: str,
        method: str,
        headers: dict[str, str],
        data: bytes | None,
        context: str,
        transport: Transport,
        timeout: float,
        deadline: float,
    ) -> tuple[int, bytes, Any]:
        """One endpoint; 429 is retried here (throttling is per account, not per host)."""
        retries = get_config().http.max_retries
        for n in range(retries + 1):
            try:
                return transport(
                    base + path,
                    method=method,
                    headers=headers_for(base, headers),
                    data=data,
                    timeout=timeout,
                    max_retries=0,
                    context=context,
                    kind="chat",
                )
            except Mcp365Error as err:
                if err.http_status != 429 or n >= retries:
                    raise
                try:
                    wait = min(float(err.retry_after or ""), 30.0)
                except ValueError:
                    wait = 0.75 * (2**n)
                if self._clock() + wait > deadline:
                    raise
                self._sleep(wait)
        raise AssertionError("unreachable")  # pragma: no cover

    @staticmethod
    def _outcome_unknown(base: str, err: Mcp365Error, context: str) -> TransportError:
        what = f"HTTP {err.http_status}" if err.http_status is not None else err.message
        out = TransportError(
            f"Teams Chat Service ({_host(base)}) đã nhận yêu cầu {context or 'ghi'} nhưng không trả kết quả rõ ràng: "
            f"{what}.",
            "Yêu cầu có thể ĐÃ được thực hiện nên KHÔNG tự gửi lại (tránh gửi trùng). Đọc lại chat để kiểm tra "
            "đã gửi/ghi được chưa rồi mới thử lại.",
        )
        out.http_status = err.http_status
        return out

    @staticmethod
    def _exhausted(failures: list[tuple[str, Mcp365Error]], context: str, replayable: bool) -> Mcp365Error:
        detail = "; ".join(f"{_host(b)}: {_short(e)}" for b, e in failures) or "hết thời gian cho phép"
        hosts = {b for b, _e in failures}
        cls = ConnectError if all(isinstance(e, ConnectError) for _b, e in failures) else TransportError
        remediation = (
            "Mạng/VPN hoặc máy chủ Microsoft đang lỗi. Chạy `check_365_connection` để xem endpoint nào còn sống, "
            "rồi thử lại sau ít phút. Có thể đổi danh sách/thứ tự endpoint bằng `[teams] chat_endpoints` hoặc "
            "MCP365_TEAMS_CHAT_ENDPOINTS."
        )
        if any("timed out" in e.message for _b, e in failures):
            remediation += " Timeout ở Chat Service gần như luôn là do máy chủ; tăng `timeout_seconds` hiếm khi giúp."
        if not replayable:
            remediation = "Yêu cầu CHƯA tới máy chủ nên chưa gửi/ghi gì. " + remediation
        out = cls(
            f"Không gọi được Teams Chat Service ({context or 'request'}) sau {len(failures)} lần thử qua "
            f"{len(hosts)} endpoint: {detail}.",
            remediation,
        )
        if failures:
            out.http_status = failures[-1][1].http_status
        return out


def _is_endpoint_fault(err: Mcp365Error) -> bool:
    """Network trouble or a server-side fault - something another front door may not have."""
    if isinstance(err, NetworkError):
        return True
    status = err.http_status
    return status is not None and (status >= 500 or status == 408)


#: The process-wide router. Looked up at call time so tests can swap it.
ROUTER = ChatServiceRouter()
