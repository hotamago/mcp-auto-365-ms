"""Teams chat-list sections and Favorites: CSA ``conversationFolders`` (Teams "new" sidebar).

Only the signed-in user's own sidebar changes; nobody else sees it or is notified, and no chat is
created, left or deleted. CSA (``chatsvcagg``) refuses the Chat Service skypetoken and the
middle-tier token (401, checked 18/11), so it gets Teams Web's own SPA token, minted from the
browser's Microsoft sign-in like Outlook mail (:class:`common.spa_auth.SpaTokenManager`).

Every write carries the newest ``folderHierarchyVersion``. A stale one is refused with 412: only
item moves and renames are retried, once, after re-reading the hierarchy and checking the plan is
unchanged; a reorder, delete or create is reported as a conflict instead (KHAO-SAT-TEAMS-NHOM §5.4).

Server limits seen live 19/11: one action per write (400 "Only drag and drop scenario supports 2
actions in payload" - the one pair it takes is ``AddItem`` then ``RemoveItem``, a drag-drop move);
at most :data:`MAX_ITEMS` chats per section (400 ``MaxItemsPerFolderThreshold``); 429 with
``Retry-After`` (~1 min) after about 100 calls in a few minutes. A 429 wrote nothing, so it is resent.
"""

from __future__ import annotations

import json
import logging
import time
import urllib.parse
from collections.abc import Callable
from dataclasses import dataclass, field
from typing import Any

from common.config import get_config
from common.errors import AuthExpiredError, ConcurrentEditError, Mcp365Error
from common.http import decode_json, request
from common.spa_auth import SpaSettings, SpaTokenManager

from .auth import TeamsAuthManager
from .client import fold

logger = logging.getLogger(__name__)

#: Teams Web adds these to every conversationFolders call; the server answers differently without them.
QUERY = {"supportsAdditionalSystemGeneratedFolders": "true", "supportsSliceItems": "true"}
#: The only folder types that take chats (Teams Web ``Oi``); every other type is system-managed.
CUSTOM_TYPES = ("UserCreated", "Favorites")
SORT_TYPES = ("MostRecent", "UnreadThenMostRecent", "UnreadOnly", "NameAlphabetical", "UserDefinedCustomOrder")
NOTE_SIDEBAR_ONLY = "Chỉ danh sách chat của bạn thay đổi; không ai khác thấy hay được báo."
#: Chats a section takes (CSA ``MaxItemsPerFolderThreshold``, hit live 19/11).
MAX_ITEMS = 50
#: Pause between the writes of one batch, to stay under CSA's rate limit.
WRITE_PACE_SECONDS = 0.5
#: 429 waits per write before that chat is reported as not done; each wait follows Retry-After.
RATE_LIMIT_WAITS = 3
DEFAULT_RETRY_AFTER = 60.0
MAX_RETRY_AFTER = 120.0


class CsaAuthManager(SpaTokenManager):
    """Teams Web's short-lived ``chatsvcagg`` token, minted from Chrome's Microsoft sign-in."""

    def _settings(self) -> SpaSettings:
        cfg = get_config().teams
        tenant = cfg.csa_tenant_id.strip() or TeamsAuthManager.get_identity().tenant_id or "organizations"
        return SpaSettings(
            product="Teams (mục chat)",
            client_id=cfg.csa_client_id,
            tenant_id=tenant,
            login_host=cfg.csa_login_host,
            origin=cfg.csa_origin,
            scope=cfg.csa_scope,
            redirect_uri=cfg.csa_redirect_uri,
            audience=cfg.csa_audience,
            open_url=cfg.csa_origin,
        )

    def _username(self) -> str:
        upn = TeamsAuthManager.get_identity().upn
        if not upn:
            raise AuthExpiredError(
                "Phiên Teams không chứa UPN để chọn tài khoản cho mục chat.",
                "Mở lại https://teams.microsoft.com trong Chrome rồi thử lại.",
            )
        return upn


def remediation() -> str:
    return (f"Mở {get_config().teams.csa_origin} trong Chrome, đăng nhập đúng tài khoản và chọn 'Stay signed in', "
            "rồi thử lại. Không cần Graph permission.")


# ----------------------------------------------------------------- hierarchy


def ordered(state: dict[str, Any]) -> list[dict[str, Any]]:
    """Live folders in display order: ``conversationFolderOrder``, then any folder it leaves out."""
    folders = {f.get("id"): f for f in state.get("conversationFolders") or [] if isinstance(f, dict) and f.get("id")}
    order = [fid for fid in state.get("conversationFolderOrder") or [] if fid in folders]
    order += [fid for fid in folders if fid not in order]
    return [folders[fid] for fid in order if not folders[fid].get("isDeleted")]


def items(folder: dict[str, Any]) -> list[str]:
    return [str(i.get("conversationId") or i.get("id") or "") for i in folder.get("conversationFolderItems") or []
            if isinstance(i, dict) and (i.get("conversationId") or i.get("id"))]


def label(folder: dict[str, Any]) -> str:
    return folder.get("name") or folder.get("folderType") or folder.get("id", "")


def find_section(state: dict[str, Any], query: str, types: tuple[str, ...] = CUSTOM_TYPES) -> dict[str, Any]:
    """One folder by ID, exact name (case/diacritics ignored) or system type name; never a partial match."""
    query = (query or "").strip()
    if not query:
        raise Mcp365Error("Chưa chọn mục.", "Truyền tên hoặc ID mục (xem list_teams_sections).")
    folders = ordered(state)
    for match in (lambda f: f.get("id") == query,
                  lambda f: fold(f.get("name") or "") == fold(query),
                  lambda f: f.get("folderType") != "UserCreated" and fold(f.get("folderType") or "") == fold(query)):
        found = [f for f in folders if match(f)]
        if len(found) > 1:
            raise Mcp365Error(f"'{query}' khớp nhiều mục: " + ", ".join(f"{label(f)} (`{f['id']}`)" for f in found),
                              "Truyền ID mục.")
        if found:
            if found[0].get("folderType") not in types:
                raise Mcp365Error(
                    f"Mục `{label(found[0])}` là mục hệ thống ({found[0].get('folderType')}); thao tác này chỉ dùng cho "
                    + " / ".join(types) + ".",
                    "Chọn mục bạn tự tạo (xem list_teams_sections).",
                )
            return found[0]
    raise Mcp365Error(f"Không có mục '{query}'.", "Xem list_teams_sections.")


def sections_of(state: dict[str, Any], chat_id: str) -> list[dict[str, Any]]:
    """Custom folders (UserCreated / Favorites) holding ``chat_id``: Teams keeps a chat in at most one."""
    key = chat_id.casefold()
    return [f for f in ordered(state) if f.get("folderType") in CUSTOM_TYPES
            and any(i.casefold() == key for i in items(f))]


def favorites(state: dict[str, Any]) -> dict[str, Any]:
    if state.get("migratedPinnedToFavorites") is not True:
        raise Mcp365Error("Tài khoản chưa chuyển chat ghim sang mục Favorites (Teams mới); ghim kiểu cũ không hỗ trợ.",
                          "Mở Teams mới trong Chrome một lần rồi thử lại.")
    found = [f for f in ordered(state) if f.get("folderType") == "Favorites"]
    if len(found) != 1:
        raise Mcp365Error(f"Teams trả {len(found)} mục Favorites, cần đúng 1.", "Xem list_teams_sections.")
    return found[0]


@dataclass(frozen=True)
class Write:
    """One CSA write, minus the version fields. ``safe``: may be retried once after a 412."""

    method: str
    folder_id: str = ""
    body: dict[str, Any] = field(default_factory=dict)
    safe: bool = False

    def key(self) -> str:
        return json.dumps([self.method, self.folder_id, self.body], sort_keys=True, ensure_ascii=False)


def move_actions(state: dict[str, Any], chat_ids: list[str], target: dict[str, Any] | None,
                 remove: bool = False) -> tuple[list[dict], list[str]]:
    """``AddItem`` / ``RemoveItem`` actions putting each chat into ``target``, or with ``remove`` taking it out
    of ``target`` (``None``: out of every custom section). One draft line per chat; a chat already where it
    should be adds no action.
    """
    actions: list[dict] = []
    lines: list[str] = []
    for chat in chat_ids:
        current = sections_of(state, chat)
        if remove:
            old = [f for f in current if target is None or f["id"] == target["id"]]
        else:
            old = [f for f in current if f["id"] != target["id"]]
            if len(old) == len(current):
                actions.append({"action": "AddItem", "folderId": target["id"], "itemId": chat})
        actions += [{"action": "RemoveItem", "folderId": f["id"], "itemId": chat} for f in old]
        moved = ", ".join(f"`{label(f)}`" for f in old)
        if remove:
            lines.append(f"bỏ khỏi mục {moved}" if old else
                         f"không ở {'mục nào' if target is None else f'`{label(target)}`'}, bỏ qua")
        elif len(old) == len(current):
            lines.append(f"vào `{label(target)}`" + (f", **chuyển** khỏi mục {moved}" if old else ""))
        else:
            lines.append(f"đã ở `{label(target)}`" + (f"; bỏ khỏi mục {moved}" if old else ", bỏ qua"))
    return actions, lines


def requests_for(actions: list[dict]) -> list[list[dict]]:
    """Split actions into CSA requests: one action each, except a leading drag-drop pair (AddItem, RemoveItem)."""
    head = 2 if (len(actions) >= 2 and actions[0].get("action") == "AddItem"
                 and actions[1].get("action") == "RemoveItem") else 0
    return ([actions[:head]] if head else []) + [[a] for a in actions[head:]]


def over_capacity(state: dict[str, Any], chat_ids: list[str], target: dict[str, Any]) -> str:
    """Why ``target`` cannot take ``chat_ids`` (``""`` when it can): a section holds at most :data:`MAX_ITEMS`."""
    have = {i.casefold() for i in items(target)}
    new = list(dict.fromkeys(c.casefold() for c in chat_ids if c.casefold() not in have))
    if len(have) + len(new) <= MAX_ITEMS:
        return ""
    return (f"Mục `{label(target)}` đang có {len(have)} chat, thêm {len(new)} thành {len(have) + len(new)}, "
            f"quá giới hạn {MAX_ITEMS} chat mỗi mục của Teams.")


def full_state(data: Any) -> bool:
    """A write response that is the whole new hierarchy, so the next write needs no GET."""
    return (isinstance(data, dict) and isinstance(data.get("conversationFolders"), list)
            and "folderHierarchyVersion" in data and not data.get("isPartialResponse"))


def _retry_after(exc: Mcp365Error) -> float:
    try:
        wait = float(exc.retry_after or DEFAULT_RETRY_AFTER)
    except (TypeError, ValueError):
        wait = DEFAULT_RETRY_AFTER
    return min(max(wait, 1.0), MAX_RETRY_AFTER)


def describe_error(exc: Mcp365Error) -> str:
    """One line per chat that was not done: the limit in plain words, else the status plus CSA's own error code."""
    text = exc.message or type(exc).__name__
    if "MaxItemsPerFolderThreshold" in text:
        return f"mục đã đủ {MAX_ITEMS} chat (giới hạn của Teams); tách bớt sang mục khác"
    if exc.http_status == 429:
        return "Teams vẫn giới hạn tần suất (HTTP 429) sau nhiều lần chờ; chưa ghi"
    first, _, rest = text.partition("\nPhản hồi: ")
    try:
        detail = json.loads(rest) if rest else {}
    except ValueError:
        detail = {}
    if isinstance(detail, dict) and (detail.get("errorCode") or detail.get("errorMessage")):
        return f"{first} {detail.get('errorCode') or ''}: {detail.get('errorMessage') or ''}".strip(" :")
    return first.splitlines()[0]


# ----------------------------------------------------------------------- HTTP


class SectionsClient:
    def __init__(self, auth: CsaAuthManager | None = None) -> None:
        self.auth = auth or CsaAuthManager()

    @staticmethod
    def url(folder_id: str = "") -> str:
        root = get_config().teams.csa_api_root.rstrip("/")
        path = "/teams/users/me/conversationFolders" + (f"/{urllib.parse.quote(folder_id, safe='')}" if folder_id else "")
        return f"{root}{path}?{urllib.parse.urlencode(QUERY)}"

    def _call(self, method: str, folder_id: str = "", body: dict | None = None, *, context: str) -> dict[str, Any]:
        url = self.url(folder_id)
        data = None if body is None else json.dumps(body, ensure_ascii=False).encode()
        for attempt in range(2):
            try:
                headers = {"Authorization": f"Bearer {self.auth.get_token(force_refresh=attempt > 0)}",
                           "Accept": "application/json"}
            except AuthExpiredError as exc:
                raise AuthExpiredError(exc.message, remediation()) from exc
            if data is not None:
                headers["Content-Type"] = "application/json"
            try:
                # A write is never resent on 5xx/timeouts: it may have landed.
                _status, raw, _headers = request(url, method=method, headers=headers, data=data, context=context,
                                                 max_retries=None if method == "GET" else 0)
                return decode_json(raw, url)
            except AuthExpiredError as exc:
                # 401 = refused before anything was applied: one fresh token, then give up.
                if attempt == 0 and exc.http_status == 401:
                    continue
                raise AuthExpiredError(exc.message, remediation()) from exc
            except Mcp365Error as exc:
                if exc.http_status == 400 and "MaxItemsPerFolderThreshold" in exc.message:
                    error = Mcp365Error(f"Mục đã đủ {MAX_ITEMS} chat, giới hạn của Teams (MaxItemsPerFolderThreshold). "
                                        "CHƯA ghi chat này.",
                                        "Tạo mục khác (manage_teams_section create) rồi xếp phần còn lại vào đó.")
                    error.http_status = 400
                    raise error from exc
                raise
        raise AssertionError("unreachable")

    def state(self) -> dict[str, Any]:
        data = self._call("GET", context="đọc mục trong danh sách chat Teams")
        if not isinstance(data.get("conversationFolders"), list) or "folderHierarchyVersion" not in data:
            raise Mcp365Error("Teams không trả danh sách mục hợp lệ.", remediation())
        return data

    def _send(self, write: Write, version: Any, context: str) -> dict[str, Any]:
        body = {"folderHierarchyVersion": version, **{k: True for k in QUERY}, **write.body}
        return self._call(write.method, write.folder_id, body, context=context)

    def apply(self, state: dict[str, Any], plan: Callable[[dict[str, Any]], Write], context: str) -> dict[str, Any]:
        """Send ``plan(state)`` with ``state``'s version. On 412: a safe plan is rebuilt from a fresh read and
        sent once more if it is unchanged; anything else is a conflict and nothing more is sent."""
        write = plan(state)
        try:
            return self._send(write, state["folderHierarchyVersion"], context)
        except Mcp365Error as exc:
            if exc.http_status != 412:
                raise
            if not write.safe:
                raise _conflict("Thao tác này không tự thử lại.") from exc
            fresh = self.state()
            again = plan(fresh)
            if again.key() != write.key():
                raise _conflict("Thao tác tính lại trên danh sách mới đã khác bản nháp.") from exc
            try:
                return self._send(again, fresh["folderHierarchyVersion"], context)
            except Mcp365Error as second:
                if second.http_status == 412:
                    raise _conflict("Lần thử lại cũng bị từ chối.") from second
                raise


    def _paced(self, call: Callable[[], dict[str, Any]], sleep: Callable[[float], None]) -> dict[str, Any]:
        """``call()``, waiting out a 429 (Retry-After) up to :data:`RATE_LIMIT_WAITS` times: a 429 wrote nothing."""
        for waits in range(RATE_LIMIT_WAITS + 1):
            try:
                return call()
            except Mcp365Error as exc:
                if exc.http_status != 429 or waits == RATE_LIMIT_WAITS:
                    raise
                wait = _retry_after(exc)
                logger.info("CSA 429: chờ %.0f s rồi gửi lại", wait)
                sleep(wait)
        raise AssertionError("unreachable")

    def apply_each(self, state: dict[str, Any], keys: list[str],
                   steps: Callable[[dict[str, Any], str], list[list[dict]]], context: str,
                   sleep: Callable[[float], None] | None = None) -> list[tuple[str, str]]:
        """For each key, send ``steps(state, key)`` one request at a time, each request re-planned on the newest
        hierarchy (taken from the previous write's response when it is complete, else read again).

        Returns ``(key, "")`` for done or ``(key, why)`` for not done; one failure does not stop the rest.
        """
        sleep = sleep or time.sleep
        results: list[tuple[str, str]] = []
        current, sent = state, 0
        for key in keys:
            try:
                last: list[dict] | None = None
                for _ in range(len(steps(current, key))):
                    todo = steps(current, key)
                    if not todo:
                        break
                    if todo[0] == last:  # never send the same request twice for one chat
                        break

                    def plan(s: dict[str, Any], key: str = key) -> Write:
                        again = steps(s, key)
                        return Write("POST", body={"actions": again[0] if again else []}, safe=True)

                    if sent:
                        sleep(WRITE_PACE_SECONDS)
                    sent += 1
                    base = current
                    reply = self._paced(lambda base=base, plan=plan: self.apply(base, plan, context), sleep)
                    last = todo[0]
                    current = reply if full_state(reply) else self._paced(self.state, sleep)
                if steps(current, key):
                    raise Mcp365Error("Teams nhận yêu cầu nhưng danh sách mục đọc lại chưa đổi; không gửi lại.",
                                      "Đọc lại list_teams_sections trước khi thử lại.")
                results.append((key, ""))
            except Mcp365Error as exc:
                results.append((key, describe_error(exc)))
                try:
                    current = self._paced(self.state, sleep)
                except Mcp365Error:
                    pass
        return results


def _conflict(why: str) -> ConcurrentEditError:
    error = ConcurrentEditError(
        f"Danh sách mục vừa thay đổi ở nơi khác (HTTP 412). {why} CHƯA ghi gì.",
        "Xem lại list_teams_sections rồi tạo bản nháp mới (is_user_confirm=false).",
    )
    error.http_status = 412
    return error
