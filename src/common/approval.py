"""User confirmation for every outbound action.

Every tool that sends, edits, deletes or overwrites something takes a
**required** ``is_user_confirm`` argument. Its schema description tells the
calling model that it may only pass ``true`` after showing the user the exact
content and receiving an explicit yes. Without it the tool refuses and hands
the draft back, so the model has something concrete to ask about.

This is deliberately the whole mechanism: nothing is blocked by destination,
nothing is queued. Microsoft Teams is sensitive, so the rule is simple and the
same everywhere - ask the user, every time, then send.
"""

from __future__ import annotations

from typing import Annotated

from pydantic import Field

from .errors import ApprovalRequiredError

CONFIRM_DESCRIPTION = (
    "BẮT BUỘC. Chỉ được đặt true SAU KHI đã hỏi ý kiến người dùng: cho họ xem nguyên văn nội dung "
    "sẽ gửi/ghi (và gửi tới đâu) và họ đã đồng ý rõ ràng cho đúng nội dung đó. Một câu chung chung "
    "như 'cứ làm đi', 'gửi luôn đi' nói TRƯỚC khi có bản nháp KHÔNG phải là đồng ý. Đồng ý cho một "
    "tin không áp dụng cho tin khác. Chưa hỏi thì đặt false để nhận bản nháp về hỏi người dùng. "
    "MUST be true only after the user has seen this exact content and explicitly approved it."
)

#: Type for the required argument. Declared once so every tool shows the same
#: rule in its schema.
UserConfirm = Annotated[bool, Field(description=CONFIRM_DESCRIPTION)]


def require_confirm(is_user_confirm: bool, action: str, target: str, preview: str) -> None:
    """Refuse unless the user confirmed; the refusal carries the draft to show them."""
    if is_user_confirm is True:
        return
    raise ApprovalRequiredError(
        f"CHƯA GỬI — {action} cần người dùng duyệt.\n\n"
        f"**Tới:** {target}\n\n**Nội dung:**\n\n{preview}",
        "Cho người dùng xem nguyên văn bản nháp trên và hỏi họ có đồng ý không. Chỉ khi họ đồng ý "
        "rõ ràng mới gọi lại tool với is_user_confirm=true.",
    )
