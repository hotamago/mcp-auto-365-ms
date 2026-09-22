"""The Outlook mail watcher's filtering, output and exit codes, verified without a network."""

from __future__ import annotations

from datetime import UTC, datetime, timedelta

import pytest

import watch_mail
from common.errors import AuthExpiredError, KeyringError, RateLimitedError, TransportError

SINCE = datetime(2026, 9, 22, 1, 0, tzinfo=UTC)


def _mail(minutes: int, sender: str = "Nguyễn Phan Nam Sơn <nam.son@example.com>", subject: str = "Hello", **kw):
    msg = {
        "id": f"AAMk-{minutes}",
        "subject": subject,
        "from": sender,
        "received": (SINCE + timedelta(minutes=minutes)).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "is_read": False,
        "has_attachments": False,
        "importance": "Normal",
        "is_flagged": False,
        "preview": "Nội dung",
    }
    msg.update(kw)
    return msg


def _ids(messages):
    return [m["id"] for m in messages]


class _Client:
    """Stands in for OutlookMailClient: each poll pops the next result (or raises it)."""

    def __init__(self, *results):
        self.results = list(results)
        self.calls = []

    def list_messages(self, **kwargs):
        self.calls.append(kwargs)
        result = self.results.pop(0) if self.results else []
        if isinstance(result, Exception):
            raise result
        return result


@pytest.fixture
def client(monkeypatch):
    holder = {}

    def install(*results):
        holder["client"] = _Client(*results)
        monkeypatch.setattr(watch_mail, "OutlookMailClient", lambda: holder["client"])
        return holder["client"]

    return install


def test_only_mail_received_at_or_after_the_cursor_counts_oldest_first():
    msgs = [_mail(5), _mail(-1), _mail(0), _mail(2)]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE)) == ["AAMk-0", "AAMk-2", "AAMk-5"]


def test_from_matches_name_or_address_without_diacritics():
    msgs = [_mail(1), _mail(2, sender="Trịnh Anh Tuấn <tuan.ta@example.com>")]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, from_names=["nam son"])) == ["AAMk-1"]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, from_names=["TUAN.TA@"])) == ["AAMk-2"]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, from_names=["nam son", "tuan"])) == ["AAMk-1", "AAMk-2"]


def test_subject_keyword_is_case_and_diacritic_insensitive():
    msgs = [_mail(1, subject="[S5] Báo cáo tuần"), _mail(2, subject="Lịch họp")]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, subjects=["bao cao"])) == ["AAMk-1"]


def test_unread_important_and_flagged_each_narrow_the_result():
    msgs = [
        _mail(1, is_read=True),
        _mail(2, importance="High"),
        _mail(3, is_flagged=True),
        _mail(4, importance="High", is_flagged=True, is_read=True),
    ]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, unread_only=True)) == ["AAMk-2", "AAMk-3"]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, important=True)) == ["AAMk-2", "AAMk-4"]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, flagged=True)) == ["AAMk-3", "AAMk-4"]
    assert _ids(watch_mail.relevant_messages(msgs, SINCE, important=True, flagged=True)) == ["AAMk-4"]


def test_render_shows_vietnam_time_sender_subject_one_line_preview_and_id():
    long_preview = "Dòng một\n\nDòng   hai " + "x" * 300
    out = watch_mail.render([_mail(25, subject="Review kiến trúc", preview=long_preview, importance="High")])
    lines = out.splitlines()
    assert lines[0] == "📧 1 mail mới"
    # 01:25 UTC is 08:25 in Vietnam.
    assert lines[1] == "- [22/09 08:25] Nguyễn Phan Nam Sơn <nam.son@example.com> · Review kiến trúc [📩 chưa đọc · ❗ quan trọng]"
    assert lines[2].startswith("  > Dòng một Dòng hai xxx") and lines[2].endswith("…")
    assert len(lines[2]) <= 4 + 160 + 1
    assert lines[3] == "  id `AAMk-25`"
    assert lines[-1].endswith("`--since 2026-09-22T01:25:01Z`")


def test_render_warns_when_the_poll_page_was_full():
    assert "Đã quét đủ 50 mail" in watch_mail.render([_mail(1)], truncated=True)
    assert "Đã quét đủ" not in watch_mail.render([_mail(1)])


def test_poll_reads_the_inbox_through_list_messages_with_a_utc_cursor():
    stub = _Client([_mail(1)] * 50)
    args = watch_mail.parse_args(["--unread-only"])
    found, truncated = watch_mail.poll_once(stub, SINCE + timedelta(microseconds=500), args)
    assert stub.calls == [
        {"folder": "inbox", "limit": 50, "unread_only": True, "since": "2026-09-22T01:00:00+00:00"}
    ]
    assert len(found) == 50 and truncated


def test_scan_is_capped_at_what_list_messages_allows():
    assert watch_mail.parse_args(["--scan", "500"]).scan == 50
    assert watch_mail.parse_args(["--scan", "0"]).scan == 1


def test_new_mail_prints_and_exits_zero(client, capsys):
    client([_mail(3, subject="Kết quả test")])
    code = watch_mail.main(["--since", "2026-09-22T01:00:00Z", "--timeout", "60"], sleep=lambda _s: None)
    out = capsys.readouterr().out
    assert code == watch_mail.EXIT_FOUND
    assert "Kết quả test" in out and "id `AAMk-3`" in out


def test_nothing_before_the_timeout_exits_three(client, capsys):
    stub = client([], [], [])
    code = watch_mail.main(
        ["--since", "2026-09-22T01:00:00Z", "--timeout", "0.05", "--interval", "0.1"], sleep=lambda _s: None
    )
    assert code == watch_mail.EXIT_TIMEOUT
    assert "Không có mail mới" in capsys.readouterr().out
    assert len(stub.calls) == 1


@pytest.mark.parametrize(
    "error",
    [
        TransportError("Request timed out after 30s.", "Thử lại."),
        RateLimitedError("HTTP 503 sau khi thử lại.", "Thử lại sau."),
    ],
)
def test_transient_errors_skip_one_poll_and_keep_watching(client, capsys, error):
    stub = client(error, [_mail(4)])
    code = watch_mail.main(["--since", "2026-09-22T01:00:00Z", "--interval", "0", "--timeout", "60"],
                           sleep=lambda _s: None)
    captured = capsys.readouterr()
    assert code == watch_mail.EXIT_FOUND
    assert len(stub.calls) == 2
    assert "bỏ qua lỗi tạm thời" in captured.err
    assert "id `AAMk-4`" in captured.out


@pytest.mark.parametrize(
    "error",
    [
        AuthExpiredError("Chrome không có phiên đăng nhập Microsoft bền để mở Outlook.", "Mở Outlook Web."),
        KeyringError("Không lấy được master key của chrome từ keyring.", "Mở khoá keyring."),
    ],
)
def test_auth_errors_stop_the_watcher_with_a_fix_on_stdout(client, capsys, error):
    client(error, [_mail(4)])
    code = watch_mail.main(["--interval", "0", "--timeout", "60"], sleep=lambda _s: None)
    out = capsys.readouterr().out
    assert code == watch_mail.EXIT_ERROR
    assert error.message in out
    assert "Stay signed in" in out


def test_invalid_since_is_rejected_before_any_request(client, capsys):
    stub = client()
    assert watch_mail.main(["--since", "hôm qua"]) == watch_mail.EXIT_ERROR
    assert stub.calls == []
