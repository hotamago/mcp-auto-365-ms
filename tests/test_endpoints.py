"""Chat Service endpoint router: ranking, failover, write safety, cooldown, config."""

from __future__ import annotations

import json
import threading
import time
import urllib.parse

import pytest

from common.config import get_config, reset_config_cache
from common.errors import AuthExpiredError, ConnectError, Mcp365Error, TransportError
from common.http import timeout_scope
from teams import endpoints
from teams.endpoints import PROBE_PATH, ChatServiceRouter, chat_bases, is_teams_media_url, split_chat_url

A = "https://teams.cloud.microsoft/api/chatsvc/apac/v1"
B = "https://teams.microsoft.com/api/chatsvc/apac/v1"
C = "https://apac.ng.msg.teams.microsoft.com/v1"
BASES = [A, B, C]
HOST_A, HOST_B, HOST_C = "teams.cloud.microsoft", "teams.microsoft.com", "apac.ng.msg.teams.microsoft.com"
HEADERS = {"Authentication": "skypetoken=t", "Origin": "https://teams.microsoft.com"}


class Clock:
    """Manual monotonic clock; sleeping advances it."""

    def __init__(self) -> None:
        self.now = 1000.0
        self.slept: list[float] = []

    def __call__(self) -> float:
        return self.now

    def sleep(self, seconds: float) -> None:
        self.slept.append(seconds)
        self.now += seconds


def http_error(status: int, retry_after: str | None = None) -> Mcp365Error:
    err = (AuthExpiredError if status == 401 else Mcp365Error)(f"HTTP {status}")
    err.http_status = status
    err.retry_after = retry_after
    return err


class FakeService:
    """Stands in for ``common.http.request``: scripted per host, records every call."""

    def __init__(self, **by_host) -> None:
        # host -> list of outcomes consumed in order (last one repeats)
        self.script = {host.replace("_", "."): list(v) if isinstance(v, list) else [v] for host, v in by_host.items()}
        self.calls: list[tuple[str, dict]] = []
        self.lock = threading.Lock()

    def __call__(self, url: str, **kwargs):
        host = urllib.parse.urlsplit(url).netloc
        with self.lock:
            self.calls.append((url, kwargs))
            queue = self.script.get(host, [(200, b"{}", {})])
            outcome = queue.pop(0) if len(queue) > 1 else queue[0]
        if callable(outcome) and not isinstance(outcome, BaseException):
            outcome = outcome(url, **kwargs)
        if isinstance(outcome, BaseException):
            raise outcome
        return outcome

    def hosts(self, *, probes: bool = False) -> list[str]:
        return [
            urllib.parse.urlsplit(url).netloc for url, _kw in self.calls if probes or not url.endswith(PROBE_PATH)
        ]


def service(**by_host) -> FakeService:
    return FakeService(**{k: v for k, v in by_host.items()})


def make_router(clock: Clock | None = None, *, probe: bool = False) -> ChatServiceRouter:
    clock = clock or Clock()
    return ChatServiceRouter(clock=clock, sleep=clock.sleep, probe_enabled=probe)


def call(router, fake, path="/users/ME/conversations?pageSize=1", method="GET", data=None):
    return router.call("apac", path, method=method, headers=HEADERS, data=data, context="test", transport=fake)


# ------------------------------------------------------------ ranking


def test_endpoints_default_to_the_three_front_doors_in_order():
    assert chat_bases("apac") == BASES


def test_fastest_endpoint_ranks_first_and_failed_probes_rank_last():
    router = make_router()
    router.record_probe(A, 0.9)
    router.record_probe(B, 0.2)
    router.record_probe(C, 4.0, ConnectError("handshake timed out"))
    assert router.rank(BASES) == [B, A, C]


def test_latency_is_smoothed_across_probe_rounds():
    """One slow sample must not flip the ranking on its own."""
    router = make_router()
    router.record_probe(A, 0.5)
    router.record_probe(B, 0.6)
    router.record_probe(A, 0.9)  # A smoothed to 0.7: now behind B (0.6)
    assert router.rank(BASES)[:2] == [B, A]
    router.record_probe(B, 4.4)  # B: 2.5 smoothed
    assert router.rank(BASES)[:2] == [A, B]


def test_endpoint_answering_the_probe_with_an_odd_status_ranks_after_healthy_ones():
    router = make_router()
    router.record_probe(A, 0.1, http_error(404))
    router.record_probe(B, 0.5)
    router.record_probe(C, 0.7)
    assert router.rank(BASES) == [B, C, A]


def test_a_401_probe_still_measures_the_host():
    """The session is the problem, not the host: every front door would say the same."""
    router = make_router()
    router.record_probe(A, 0.4, http_error(401))
    router.record_probe(B, 0.2, http_error(401))
    router.record_probe(C, 0.3, http_error(401))
    assert router.rank(BASES) == [B, C, A]


def test_probing_is_lazy_parallel_and_uses_the_first_answer():
    def slow(url, **kw):
        time.sleep(0.4)
        return 200, b"{}", {}

    def fast(url, **kw):
        time.sleep(0.02)
        return 200, b"{}", {}

    fake = service(teams_cloud_microsoft=slow, teams_microsoft_com=fast, **{HOST_C: ConnectError("TLS stalled")})
    router = ChatServiceRouter(probe_enabled=True)
    assert fake.calls == []  # constructing the router never touches the network

    started = time.monotonic()
    order = router.ordered(BASES, HEADERS, fake)
    assert time.monotonic() - started < 0.35  # did not wait for the slow endpoint
    assert order[0] == B

    assert router._rounds[tuple(BASES)].done.wait(2)
    assert router.rank(BASES) == [B, A, C]
    probes = [kw for url, kw in fake.calls if url.endswith(PROBE_PATH)]
    assert len(probes) == 3
    assert all(kw["timeout"] == get_config().teams.chat_probe_timeout and kw["max_retries"] == 0 for kw in probes)


# ------------------------------------------------------------ failover


def test_get_fails_over_to_the_next_endpoint_and_demotes_the_failed_one():
    fake = service(teams_cloud_microsoft=TransportError("RemoteDisconnected"), teams_microsoft_com=(200, b'{"ok":1}', {}))
    router = make_router()
    status, body, _ = call(router, fake)
    assert (status, body) == (200, b'{"ok":1}')
    assert fake.hosts() == [HOST_A, HOST_B]
    # The failed endpoint is demoted: the next request starts elsewhere.
    assert router.rank(BASES)[-1] == A
    call(router, fake)
    assert fake.hosts()[2] == HOST_B


def test_get_fails_over_on_a_gateway_error():
    fake = service(teams_cloud_microsoft=http_error(502))
    router = make_router()
    call(router, fake)
    assert fake.hosts() == [HOST_A, HOST_B]


def test_write_switches_endpoint_when_the_request_never_left():
    fake = service(teams_cloud_microsoft=ConnectError("connection refused"), teams_microsoft_com=(201, b"{}", {}))
    router = make_router()
    status, _body, _ = call(router, fake, "/users/ME/conversations/x/messages", method="POST", data=b'{"m":1}')
    assert status == 201
    assert fake.hosts() == [HOST_A, HOST_B]
    assert [kw["data"] for _u, kw in fake.calls] == [b'{"m":1}', b'{"m":1}']


@pytest.mark.parametrize(
    "failure",
    [TransportError("Remote end closed connection without response"), http_error(502), http_error(500)],
)
@pytest.mark.parametrize("method", ["POST", "PUT", "DELETE"])
def test_write_is_never_resent_once_it_may_have_landed(failure, method):
    fake = service(teams_cloud_microsoft=failure)
    router = make_router()
    with pytest.raises(TransportError) as excinfo:
        call(router, fake, "/users/ME/conversations/x/messages", method=method, data=b"{}")
    assert fake.hosts() == [HOST_A]  # nothing went to another endpoint
    assert "KHÔNG tự gửi lại" in excinfo.value.remediation
    assert "kiểm tra" in excinfo.value.remediation.lower()


@pytest.mark.parametrize("status", [400, 401, 403, 404, 409, 412])
def test_business_errors_are_returned_without_failover(status):
    fake = service(teams_cloud_microsoft=http_error(status))
    router = make_router()
    with pytest.raises(Mcp365Error) as excinfo:
        call(router, fake)
    assert excinfo.value.http_status == status
    assert fake.hosts() == [HOST_A]
    assert router.rank(BASES)[0] == A  # a business error says nothing about the host


def test_throttling_is_retried_on_the_same_endpoint():
    clock = Clock()
    fake = service(teams_cloud_microsoft=[http_error(429, retry_after="2"), (200, b"{}", {})])
    router = make_router(clock)
    call(router, fake, method="POST", data=b"{}")  # 429 = rejected, safe to resend even for a write
    assert fake.hosts() == [HOST_A, HOST_A]
    assert clock.slept == [2.0]


def test_all_endpoints_failing_gives_one_actionable_error():
    fake = service(**{h: ConnectError("unreachable") for h in (HOST_A, HOST_B, HOST_C)})
    router = make_router()
    with pytest.raises(ConnectError) as excinfo:
        call(router, fake, method="POST", data=b"{}")
    assert fake.hosts() == [HOST_A, HOST_B, HOST_C]
    assert "3 endpoint" in excinfo.value.message
    assert "CHƯA tới máy chủ" in excinfo.value.remediation


# ------------------------------------------------------------ timeouts


def test_each_attempt_gets_the_chat_timeout_within_the_total_budget():
    clock = Clock()

    def hang(url, timeout, **kw):
        clock.now += timeout  # the attempt used its whole timeout
        raise TransportError("timed out")

    fake = service(**{h: hang for h in (HOST_A, HOST_B, HOST_C)})
    router = make_router(clock)
    started = clock.now
    with pytest.raises(TransportError):
        call(router, fake)
    assert [kw["timeout"] for _u, kw in fake.calls] == [10.0, 10.0, 5.0]
    assert clock.now - started <= get_config().teams.chat_budget


def test_a_tool_timeout_applies_to_every_attempt_on_every_endpoint():
    clock = Clock()

    def hang(url, timeout, **kw):
        clock.now += timeout
        raise TransportError("timed out")

    fake = service(**{h: hang for h in (HOST_A, HOST_B, HOST_C)})
    router = make_router(clock)
    with timeout_scope(40), pytest.raises(TransportError):
        call(router, fake)
    assert [kw["timeout"] for _u, kw in fake.calls] == [40.0, 40.0, 40.0]


# ------------------------------------------------------------ cooldown


def test_failed_endpoint_cools_down_then_is_measured_again():
    clock = Clock()
    gate = threading.Event()  # holds the re-probe until the demoted order has been checked

    def answer(url, **kw):
        if len(fake.calls) > 3:
            gate.wait(2)
        return 200, b"{}", {}

    fake = service(**{h: answer for h in (HOST_A, HOST_B, HOST_C)})
    router = ChatServiceRouter(clock=clock, sleep=clock.sleep, probe_enabled=True)
    key = tuple(BASES)

    router.ordered(BASES, HEADERS, fake)
    assert router._rounds[key].done.wait(2)
    assert router.rank(BASES) == [A, B, C]

    router.mark_failed(A, TransportError("reset"), BASES)
    assert router.rank(BASES) == [B, C, A]

    clock.now += 100  # still cooling: no re-probe, still demoted
    assert router.ordered(BASES, HEADERS, fake) == [B, C, A]
    assert len(fake.hosts(probes=True)) == 3

    clock.now += 100  # cooldown (180 s) over: re-measured in the background
    assert router.ordered(BASES, HEADERS, fake)[-1] == A  # not trusted until measured
    gate.set()
    assert router._rounds[key].done.wait(2)
    assert len(fake.hosts(probes=True)) == 6
    assert router.rank(BASES)[0] == A


def test_ranking_is_refreshed_after_the_ttl():
    clock = Clock()
    fake = service()
    router = ChatServiceRouter(clock=clock, sleep=clock.sleep, probe_enabled=True)
    router.ordered(BASES, HEADERS, fake)
    assert router._rounds[tuple(BASES)].done.wait(2)
    clock.now += get_config().teams.chat_probe_ttl + 1
    router.ordered(BASES, HEADERS, fake)
    assert router._rounds[tuple(BASES)].done.wait(2)
    assert len(fake.hosts(probes=True)) == 6


# ------------------------------------------------------------ configuration


def test_endpoint_list_and_order_come_from_env(monkeypatch):
    monkeypatch.setenv(
        "MCP365_TEAMS_CHAT_ENDPOINTS",
        "https://{region}.ng.msg.teams.microsoft.com/v1, https://teams.microsoft.com/api/chatsvc/{region}/v1",
    )
    monkeypatch.setenv("MCP365_TEAMS_CHAT_PROBE", "0")
    monkeypatch.setenv("MCP365_HTTP_TIMEOUT_CHAT", "6")
    reset_config_cache()
    assert chat_bases("apac") == [C, B]
    assert get_config().teams.chat_probe is False

    fake = service(**{HOST_C: ConnectError("down")})
    router = make_router()
    call(router, fake)
    assert fake.hosts() == [HOST_C, HOST_B]
    assert fake.calls[0][1]["timeout"] == 6.0


def test_endpoint_list_comes_from_the_config_file(monkeypatch, tmp_path):
    cfg = tmp_path / "config.toml"
    cfg.write_text(
        '[teams]\nchat_endpoints = ["https://teams.microsoft.com/api/chatsvc/{region}/v1"]\nchat_cooldown = 60.0\n'
        "[http]\ntimeout_chat = 7.5\n",
        encoding="utf-8",
    )
    monkeypatch.setattr("common.config._USER_CONFIG", cfg)
    reset_config_cache()
    assert chat_bases("emea") == ["https://teams.microsoft.com/api/chatsvc/emea/v1"]
    assert get_config().teams.chat_cooldown == 60.0
    assert get_config().http.timeout_chat == 7.5


def test_unusable_endpoint_config_falls_back_to_the_defaults(monkeypatch):
    monkeypatch.setenv("MCP365_TEAMS_CHAT_ENDPOINTS", "not-a-url")
    reset_config_cache()
    assert chat_bases("apac") == BASES


# ------------------------------------------------------------ absolute links


def test_paging_link_on_the_legacy_host_goes_through_the_preferred_endpoint():
    link = f"{C}/users/ME/conversations/19%3Aabc%40thread.v2/messages?startTime=1&syncState=s1&pageSize=50"
    assert split_chat_url(link) == ("apac", "/users/ME/conversations/19%3Aabc%40thread.v2/messages?startTime=1&syncState=s1&pageSize=50")

    router = make_router()
    router.record_probe(B, 0.1)
    router.record_probe(A, 0.3)
    router.mark_failed(C, ConnectError("TLS stalled"), BASES)
    fake = service()
    call(router, fake, link)
    assert fake.calls[0][0] == f"{B}/users/ME/conversations/19%3Aabc%40thread.v2/messages?startTime=1&syncState=s1&pageSize=50"


def test_link_keeps_its_own_region():
    link = "https://teams.cloud.microsoft/api/chatsvc/emea/v1/users/ME/conversations?pageSize=5"
    fake = service(teams_cloud_microsoft=ConnectError("down"))
    call(make_router(), fake, link)
    assert fake.calls[1][0] == "https://teams.microsoft.com/api/chatsvc/emea/v1/users/ME/conversations?pageSize=5"


def test_foreign_absolute_url_is_fetched_as_is():
    fake = service()
    call(make_router(), fake, "https://example.com/x?y=1")
    assert [u for u, _kw in fake.calls] == ["https://example.com/x?y=1"]


def test_proxies_get_their_own_origin():
    fake = service()
    call(make_router(), fake)
    assert fake.calls[0][1]["headers"]["Origin"] == "https://teams.cloud.microsoft"
    fake = service(teams_cloud_microsoft=ConnectError("down"), teams_microsoft_com=ConnectError("down"))
    call(make_router(), fake)
    assert fake.calls[2][1]["headers"]["Origin"] == "https://teams.microsoft.com"  # legacy host: unchanged


@pytest.mark.parametrize(
    ("url", "expected"),
    [
        ("https://as-api.asm.skype.com/v1/objects/0-sa-d1-abc/views/imgo", True),
        ("https://apac.asyncgw.teams.microsoft.com/v1/objects/0-sa-d1-abc/views/imgo", True),
        ("https://apac.ng.msg.teams.microsoft.com/v1/objects/x/views/imgo", True),
        ("https://teams.cloud.microsoft/api/chatsvc/apac/v1/objects/x/views/imgo", True),
        ("https://teams.cloud.microsoft/api/asyncgw/apac/v1/objects/x/views/imgo", True),
        ("https://teams.microsoft.com/api/asm/v1/objects/x/views/imgo", True),
        ("https://teams.microsoft.com/l/message/19:abc/123", False),
        ("https://evil.example/?next=asm.skype.com", False),
        ("https://vingroupjsc.sharepoint.com/sites/X/a.png", False),
    ],
)
def test_teams_media_urls_are_recognised_by_host(url, expected):
    assert is_teams_media_url(url) is expected


# ------------------------------------------------------------ TeamsClient wiring


@pytest.fixture
def wired_client(monkeypatch, identity):
    from teams.client import TeamsClient

    client = TeamsClient()
    monkeypatch.setattr(client, "_auth", lambda: {"region": "apac", "token": "tok", "identity": identity})
    monkeypatch.setattr(type(client), "identity", property(lambda self: identity))
    return client


def test_teams_client_reads_fail_over(wired_client, monkeypatch):
    listing = {"conversations": [{"id": "19:abc@thread.v2", "threadProperties": {"topic": "Dev team"}}]}
    fake = service(
        teams_cloud_microsoft=ConnectError("down"),
        teams_microsoft_com=(200, json.dumps(listing).encode(), {}),
    )
    monkeypatch.setattr("teams.client.request", fake)
    chats = wired_client.list_conversations(use_cache=False)
    assert [c["name"] for c in chats] == ["Dev team"]
    assert fake.hosts() == [HOST_A, HOST_B]


def test_teams_client_send_is_not_duplicated(wired_client, monkeypatch):
    fake = service(teams_cloud_microsoft=TransportError("Remote end closed connection without response"))
    monkeypatch.setattr("teams.client.request", fake)
    with pytest.raises(TransportError):
        wired_client.send_message("19:abc@thread.v2", "xin chào")
    posts = [u for u, kw in fake.calls if kw["method"] == "POST"]
    assert len(posts) == 1


def test_image_on_a_chat_host_is_downloaded_through_the_router(wired_client, monkeypatch, tmp_path):
    fake = service(teams_cloud_microsoft=(200, b"PNG", {}))
    monkeypatch.setattr("teams.client.request", fake)
    out = wired_client.download_image(f"{C}/objects/0-x/views/imgo", tmp_path / "a.png")
    assert out.read_bytes() == b"PNG"
    url, kw = fake.calls[0]
    assert url == f"{A}/objects/0-x/views/imgo"
    assert kw["headers"]["Cookie"] == "skypetoken_asm=tok"


def test_global_router_is_swappable():
    """TeamsClient looks the router up at call time (tests and tools share one per process)."""
    assert isinstance(endpoints.ROUTER, ChatServiceRouter)
