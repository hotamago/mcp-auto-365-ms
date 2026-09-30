"""Một-hoặc-nhiều: gom tham số thành danh sách và chạy song song có giới hạn, giữ thứ tự.

Quy ước chung cho mọi tool (README "One or many"):

- Tham số số ít (``file_path``, ``local_file_path``, ``message_id``, ``chat_name_or_id``) nhận
  **một** giá trị như trước.
- Tham số số nhiều (``file_paths``, ``local_file_paths``, ``attachments``, ``message_ids``,
  ``chat_names_or_ids``) nhận danh sách; một chuỗi đơn, hoặc chuỗi JSON dạng ``["a", "b"]``
  (một số harness gửi mảng thành chuỗi), cũng được.
- Đưa cả hai thì gộp: số ít đứng đầu, rồi tới danh sách, bỏ trùng, giữ thứ tự.
"""

from __future__ import annotations

import contextvars
import json
from collections.abc import Callable, Hashable, Iterable, Sequence
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
from typing import Any

from common.config import get_config
from common.errors import Mcp365Error


def _expand(value: Any) -> list[str]:
    if value is None:
        return []
    if isinstance(value, str):
        text = value.strip()
        if text.startswith("[") and text.endswith("]"):
            try:
                parsed = json.loads(text)
            except ValueError:
                parsed = None
            if isinstance(parsed, list):
                return [str(v).strip() for v in parsed if v is not None and str(v).strip()]
        return [text] if text else []
    if isinstance(value, Iterable):
        out: list[str] = []
        for item in value:
            out.extend(_expand(item))
        return out
    text = str(value).strip()
    return [text] if text else []


def as_list(*values: Any) -> list[str]:
    """Gộp các tham số một-hoặc-nhiều thành một danh sách chuỗi: bỏ rỗng, bỏ trùng, giữ thứ tự."""
    out: list[str] = []
    for value in values:
        out.extend(_expand(value))
    return list(dict.fromkeys(out))


def local_files(*values: Any, what: str = "file đính kèm") -> list[Path]:
    """Đường dẫn file cục bộ (đã resolve) theo đúng thứ tự đưa vào; trùng file thì giữ lần đầu.

    File nào không có thì báo **tất cả** file thiếu trong một lỗi, trước khi làm gì khác.
    """
    raws = as_list(*values)
    paths: dict[Path, None] = {}
    missing: list[str] = []
    for raw in raws:
        path = Path(raw).expanduser().resolve()
        if not path.is_file():
            missing.append(raw)
        else:
            paths.setdefault(path, None)
    if missing:
        if len(missing) == 1:
            message = f"Không tìm thấy {what}: {missing[0]}"
        else:
            message = f"Không tìm thấy {len(missing)} {what}:\n" + "\n".join(f"- {m}" for m in missing)
        raise Mcp365Error(message, "Kiểm tra lại đường dẫn (dùng đường dẫn tuyệt đối nếu cần).")
    return list(paths)


def same_names(paths: Sequence[Path | str]) -> list[str]:
    """Tên file trùng nhau (không phân biệt hoa thường, như SharePoint) giữa các đường dẫn khác nhau."""
    seen: dict[str, str] = {}
    dupes: list[str] = []
    for p in paths:
        name = Path(p).name
        key = name.casefold()
        if key in seen and name not in dupes:
            dupes.append(name)
        seen.setdefault(key, name)
    return dupes


def transfer_workers() -> int:
    """Số file tải lên/tải về cùng lúc (``http.transfer_workers``, mặc định 4; 1 = tuần tự)."""
    return max(1, int(get_config().http.transfer_workers or 1))


def run_ordered[T, R](
    fn: Callable[[T], R],
    items: Sequence[T],
    *,
    key: Callable[[T], Hashable] | None = None,
    max_workers: int | None = None,
) -> list[tuple[R | None, BaseException | None]]:
    """``fn(item)`` cho mọi item, song song tối đa ``max_workers``; kết quả ``(giá trị, lỗi)`` đúng thứ tự vào.

    - Lỗi của một item được giữ lại cho item đó, không dừng các item khác: người gọi quyết
      định báo gì (chưa gửi tin nếu thiếu file, hay báo một phần).
    - Các item cùng ``key`` (vd cùng tên file đích) chạy **nối tiếp** trong một luồng, theo
      thứ tự vào: hai lượt ghi cùng một file không bao giờ chạy cùng lúc.
    - Luồng con mang theo ``contextvars`` của người gọi (``timeout_seconds`` của tool).
    - Một item (hoặc ``max_workers`` = 1): chạy ngay trên luồng gọi, y như vòng lặp cũ.
    """
    workers = transfer_workers() if max_workers is None else max(1, max_workers)
    results: list[tuple[R | None, BaseException | None]] = [(None, None)] * len(items)

    def one(index: int) -> None:
        try:
            results[index] = (fn(items[index]), None)
        except Exception as exc:  # noqa: BLE001 - trả lại cho người gọi, gắn đúng item
            results[index] = (None, exc)

    lanes: dict[Hashable, list[int]] = {}
    for index, item in enumerate(items):
        lanes.setdefault(key(item) if key else index, []).append(index)

    if len(items) <= 1 or workers == 1 or len(lanes) == 1:
        for index in range(len(items)):
            one(index)
        return results

    parent = contextvars.copy_context()

    def lane(indices: list[int]) -> None:
        for index in indices:
            one(index)

    with ThreadPoolExecutor(max_workers=min(workers, len(lanes)), thread_name_prefix="mcp365-xfer") as pool:
        # Một Context không vào được hai luồng cùng lúc: mỗi làn một bản sao.
        list(pool.map(lambda indices: parent.copy().run(lane, indices), lanes.values()))
    return results
