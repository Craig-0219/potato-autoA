"""Windows UI Automation helpers for interacting with LINE."""
from __future__ import annotations

import time
from dataclasses import dataclass
from typing import Callable, List, Optional


@dataclass
class CycleResult:
    """Summary of the friend chat cycling test."""

    processed: List[str]
    failed: List[str]
    reached_end: bool


class LineAutomationError(RuntimeError):
    """Raised when LINE automation cannot be performed."""


def cycle_friend_chats(
    *,
    limit: int,
    message: str | None,
    log: Callable[[str], None],
    pause: float = 0.35,
) -> CycleResult:
    """Cycle through LINE friend list items using UI Automation.

    Parameters
    ----------
    limit:
        Maximum number of friends to open.
    message:
        Optional text message to send after each chat is opened. If empty, the
        test only opens the chat window.
    log:
        Callback used to append log messages to the UI.
    pause:
        Delay (seconds) between UI actions to let the LINE interface update.
    """

    try:
        from pywinauto import Application
        from pywinauto.findwindows import ElementNotFoundError
        try:
            from pywinauto.findwindows import TimeoutError  # type: ignore[attr-defined]
        except ImportError:
            from pywinauto.timings import TimeoutError
    except ImportError as exc:  # pragma: no cover - dependency guaranteed at runtime
        raise LineAutomationError("載入 pywinauto 失敗，請確認已安裝並可用：" + str(exc)) from exc

    try:
        app = Application(backend="uia").connect(path="LINE.exe", timeout=3)
    except (ElementNotFoundError, TimeoutError, RuntimeError) as exc:
        raise LineAutomationError(f"無法連線到 LINE：{exc}") from exc

    window = app.window(title_re="^LINE")
    if not window.exists():
        raise LineAutomationError("找不到 LINE 主視窗，請先開啟 LINE。")

    try:
        window.set_focus()
    except Exception:
        pass

    friend_list = _locate_friend_list(window)
    if friend_list is None:
        raise LineAutomationError("找不到好友清單清單控制項。")

    list_wrapper = friend_list.wrapper_object()
    scroll = getattr(list_wrapper, "iface_scroll_pattern", None)

    processed: list[str] = []
    failed: list[str] = []
    seen_names: set[str] = set()
    visited_positions: set[float] = set()
    consecutive_no_new = 0

    def current_vertical() -> Optional[float]:
        if scroll is None:
            return None
        try:
            return float(scroll.CurrentVerticalScrollPercent)
        except Exception:
            return None

    reached_end = False

    while len(processed) < limit:
        visible_items = _collect_visible_items(list_wrapper)
        visible_items = [(name, item) for (name, item) in visible_items]
        new_item_found = False
        for name, wrapper in visible_items:
            if name in seen_names:
                continue
            seen_names.add(name)
            new_item_found = True

            if len(processed) >= limit:
                break

            log(f"聊天測試：開啟 {name}")
            success = _open_chat(wrapper, pause)
            if success and message:
                sent = _send_message(window, message, pause)
                if not sent:
                    log(f"{name}：找不到訊息輸入框，略過發送。")
            if success:
                processed.append(name)
            else:
                failed.append(name)

            if len(processed) >= limit:
                break

        if len(processed) >= limit:
            break

        if not visible_items:
            break

        if not new_item_found:
            consecutive_no_new += 1
        else:
            consecutive_no_new = 0

        if scroll is None:
            break

        previous = current_vertical()
        if previous is not None:
            visited_positions.add(previous)
            if previous >= 100.0:
                reached_end = True
                break

        try:
            list_wrapper.scroll(direction="down", amount="page")
            time.sleep(pause)
        except Exception as exc:
            log(f"好友清單捲動失敗：{exc}")
            break

        current = current_vertical()
        if current is not None:
            if current >= 100.0:
                reached_end = True
                break
            if previous is not None and current <= previous:
                consecutive_no_new += 1
            if current in visited_positions:
                consecutive_no_new += 1
        if consecutive_no_new >= 3:
            break

    return CycleResult(processed=processed, failed=failed, reached_end=reached_end)


def _locate_friend_list(window) -> Optional[object]:
    """Try to locate the LINE friend list UIA control."""

    candidates = window.descendants(control_type="List")
    if not candidates:
        return None

    # Prefer list controls whose name hints at contacts/friends.
    for candidate in candidates:
        name = (candidate.window_text() or "").strip()
        if name and any(keyword in name for keyword in ("好友", "朋友", "friend")):
            return candidate

    # Otherwise return the first list with list-item children.
    for candidate in candidates:
        try:
            if candidate.children(control_type="ListItem"):
                return candidate
        except Exception:
            continue

    return candidates[0]


def _collect_visible_items(list_wrapper) -> list[tuple[str, object]]:
    """Collect currently visible list items and their wrappers."""

    items: list[tuple[str, object]] = []
    try:
        children = list_wrapper.children()
    except Exception:
        children = []

    if not children:
        try:
            children = list_wrapper.descendants(control_type="ListItem")
        except Exception:
            children = []

    for child in children:
        try:
            wrapper = child.wrapper_object()
        except Exception:
            continue

        name = (wrapper.window_text() or "").strip()
        if not name:
            try:
                name = getattr(wrapper.element_info, "name", "").strip()
            except Exception:
                name = ""
        if not name:
            continue
        items.append((name, wrapper))
    return items


def _open_chat(item_wrapper, pause: float) -> bool:
    try:
        try:
            item_wrapper.scroll_into_view()
        except Exception:
            pass
        item_wrapper.click_input()
        time.sleep(pause)
        return True
    except Exception:
        return False


def _send_message(window, message: str, pause: float) -> bool:
    edits = window.descendants(control_type="Edit")
    target = None
    for edit in edits:
        try:
            wrapper = edit.wrapper_object()
        except Exception:
            continue
        try:
            if wrapper.is_keyboard_focusable():
                target = wrapper
                break
        except Exception:
            continue

    if target is None:
        return False

    try:
        target.click_input()
        time.sleep(0.1)
        target.type_keys("^a{BACKSPACE}", set_foreground=True)
        if message:
            target.type_keys(message, with_spaces=True, pause=0.02, set_foreground=True)
            time.sleep(0.1)
        target.type_keys("{ENTER}", set_foreground=True)
        time.sleep(pause)
    except Exception:
        return False

    return True

