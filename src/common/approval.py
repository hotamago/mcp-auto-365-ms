"""Human-approval gate for outbound actions.

Nothing in this file talks to Microsoft. It sits between a tool and its client
call and answers one question: *is this allowed to leave the machine yet?*

Three layers, weakest to strongest:

1. **Staging.** A sensitive tool does not act; it returns a rendered draft plus
   a one-time token and stores the real call. The draft lands in the transcript,
   where the human can read the exact bytes before anything is sent.
2. **Confirmation.** ``confirm_pending_action(token)`` runs the stored call.
   Tokens are single-use and expire.
3. **Destination policy.** Group chats, channels and meeting chats are refused
   outright and never get a token. This is the only layer an agent cannot talk
   its way past, because ``allow_group_sends`` lives in the user's config or
   environment - not in a tool argument the agent fills in.

The ordering matters. Layers 1-2 are procedural and assume good faith; layer 3
is the actual boundary. A real incident drove this: an agent was told "just send
it" and posted an unreviewed message into a company-wide squad channel.
"""

from __future__ import annotations

import secrets
import threading
import time
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

from .config import get_config
from .errors import ApprovalRequiredError, PendingActionError

#: Conversation types that reach more than one other person.
_MULTI_PERSON = {"groupchat", "channel", "meetingchat"}

_SELF_IDS = {"48:notes"}


@dataclass
class _Pending:
    token: str
    action: str
    target: str
    preview: str
    executor: Callable[[], Any]
    created_at: float = field(default_factory=time.time)


_PENDING: dict[str, _Pending] = {}
_LOCK = threading.Lock()


def _purge_expired() -> None:
    ttl = get_config().safety.pending_ttl_s
    now = time.time()
    for token in [t for t, p in _PENDING.items() if now - p.created_at > ttl]:
        _PENDING.pop(token, None)


def is_self_chat(conversation_id: str) -> bool:
    return (conversation_id or "").strip().lower() in _SELF_IDS


def is_multi_person(conversation: dict[str, Any]) -> bool:
    """True when the destination is seen by more than one other person.

    Falls back to inspecting the id when the type is unknown, because
    ``find_conversation`` returns ``type="Unknown"`` for a raw id that is not in
    the cache - and an unknown destination must not be treated as a safe one.
    """
    ctype = str(conversation.get("type", "")).strip().lower()
    if ctype in _MULTI_PERSON:
        return True
    if ctype == "directchat":
        return False
    conv_id = str(conversation.get("id", ""))
    return "@thread." in conv_id or conv_id.startswith("19:meeting_")


def guard_destination(conversation: dict[str, Any], action: str) -> None:
    """Refuse multi-person destinations unless the human opted in.

    Raised before staging, so a blocked destination never even produces a token.
    """
    if not is_multi_person(conversation):
        return
    if get_config().safety.allow_group_sends:
        return
    name = conversation.get("name") or conversation.get("id", "?")
    raise ApprovalRequiredError(
        f"Từ chối {action} tới '{name}' — đây là nhóm/kênh nhiều người, "
        f"mặc định bị chặn để tránh gửi nhầm ra kênh công ty.",
        "Nếu bạn thực sự muốn gửi: bật `allow_group_sends = true` trong mục [safety] của "
        "~/.config/mcp-auto-365-ms/config.toml, hoặc chạy lại server với "
        "MCP365_ALLOW_GROUP_SENDS=1. Agent không được tự đặt cờ này.",
    )


def stage(action: str, target: str, preview: str, executor: Callable[[], Any]) -> str:
    """Hold an action as a draft and return it for human review.

    When approval is switched off, or the target is the personal notes chat, the
    executor runs immediately and its result is returned unchanged.
    """
    cfg = get_config().safety
    if not cfg.require_approval:
        return str(executor())

    with _LOCK:
        _purge_expired()
        token = secrets.token_hex(4)
        _PENDING[token] = _Pending(token=token, action=action, target=target, preview=preview, executor=executor)

    minutes = int(cfg.pending_ttl_s // 60)
    return (
        f"# ⏸️ Nháp — CHƯA GỬI\n\n"
        f"| | |\n| --- | --- |\n"
        f"| **Hành động** | {action} |\n"
        f"| **Tới** | {target} |\n"
        f"| **Mã duyệt** | `{token}` |\n\n"
        f"## Nội dung sẽ gửi\n\n{preview}\n\n"
        f"---\n"
        f"**Chưa có gì được gửi đi.** Đưa bản nháp này cho người dùng xem và chờ họ đồng ý.\n"
        f"Khi được đồng ý: `confirm_pending_action(\"{token}\")` · Bỏ: `cancel_pending_action(\"{token}\")`\n"
        f"Mã hết hạn sau {minutes} phút."
    )


def confirm(token: str) -> Any:
    """Run a staged action exactly once."""
    key = (token or "").strip().strip("`\"'")
    with _LOCK:
        _purge_expired()
        item = _PENDING.pop(key, None)
    if item is None:
        raise PendingActionError(
            f"Không tìm thấy hành động đang chờ với mã '{token}'.",
            "Mã có thể đã dùng rồi, đã huỷ, hoặc đã hết hạn. Gọi `list_pending_actions` để xem "
            "danh sách còn hiệu lực, hoặc soạn lại bản nháp.",
        )
    return item.executor()


def cancel(token: str) -> str:
    key = (token or "").strip().strip("`\"'")
    with _LOCK:
        item = _PENDING.pop(key, None)
    if item is None:
        raise PendingActionError(
            f"Không có hành động đang chờ nào mang mã '{token}'.",
            "Gọi `list_pending_actions` để xem danh sách.",
        )
    return f"✓ Đã huỷ bản nháp `{key}` ({item.action} → {item.target}). Không có gì được gửi."


def pending_rows() -> list[dict[str, Any]]:
    with _LOCK:
        _purge_expired()
        items = list(_PENDING.values())
    ttl = get_config().safety.pending_ttl_s
    now = time.time()
    return [
        {
            "token": p.token,
            "action": p.action,
            "target": p.target,
            "preview": p.preview,
            "expires_in_s": max(0, int(ttl - (now - p.created_at))),
        }
        for p in sorted(items, key=lambda x: x.created_at)
    ]


def reset() -> None:
    """Testing hook: drop every staged action."""
    with _LOCK:
        _PENDING.clear()
