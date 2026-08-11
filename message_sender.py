from __future__ import annotations

import logging
import time
from typing import Callable, Optional

import psutil
import pyautogui
import pyperclip
import win32api
import win32con
import win32gui
import win32process

from config import Config

logger = logging.getLogger("kbond_watcher")


class MessageSenderError(RuntimeError):
    pass


def relative_point(
    left: int,
    top: int,
    width: int,
    height: int,
    ratio_x: float,
    ratio_y: float,
) -> tuple[int, int]:
    return left + int(width * ratio_x), top + int(height * ratio_y)


def _process_pids(process_name: str) -> list[int]:
    expected = process_name.lower()
    pids: list[int] = []
    for proc in psutil.process_iter(["name", "pid"]):
        try:
            if (proc.info.get("name") or "").lower() == expected:
                pids.append(int(proc.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids


def _is_process_running(process_name: str) -> bool:
    return bool(_process_pids(process_name))


def _rank_window(hwnd: int) -> int:
    rank = 0
    try:
        if win32gui.IsWindowVisible(hwnd):
            rank += 4
        if not win32gui.IsIconic(hwnd):
            rank += 2
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        rank += max(0, (right - left) * (bottom - top)) // 10_000
    except Exception:
        pass
    return rank


def _find_best(matcher: Callable[[int], bool]) -> Optional[int]:
    matches: list[int] = []

    def _callback(hwnd: int, _: object) -> bool:
        try:
            if matcher(hwnd):
                matches.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_callback, None)
    if not matches:
        return None
    return max(matches, key=_rank_window)


def _is_main_window(hwnd: int, cfg: Config, pids: set[int]) -> bool:
    if win32gui.GetClassName(hwnd) != cfg.kakao_window_class:
        return False
    if not win32gui.GetWindowText(hwnd).startswith(cfg.kakao_main_title):
        return False
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return int(pid) in pids


def _is_room_window(hwnd: int, cfg: Config, room_name: str, pids: set[int]) -> bool:
    if win32gui.GetClassName(hwnd) != cfg.kakao_window_class:
        return False
    title = win32gui.GetWindowText(hwnd)
    if not title or title.startswith(cfg.kakao_main_title):
        return False
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    if int(pid) not in pids:
        return False
    return room_name in title


def find_main_window(cfg: Config) -> Optional[int]:
    pids = set(_process_pids(cfg.kakao_process_name))
    if not pids:
        return None
    return _find_best(lambda hwnd: _is_main_window(hwnd, cfg, pids))


def find_room_window(cfg: Config, room_name: str) -> Optional[int]:
    pids = set(_process_pids(cfg.kakao_process_name))
    if not pids:
        return None
    return _find_best(lambda hwnd: _is_room_window(hwnd, cfg, room_name, pids))


def _wait_until(
    predicate: Callable[[], Optional[int]],
    timeout_seconds: float,
    poll_seconds: float,
) -> Optional[int]:
    deadline = time.monotonic() + max(timeout_seconds, 0)
    while time.monotonic() < deadline:
        hwnd = predicate()
        if hwnd:
            return hwnd
        time.sleep(max(poll_seconds, 0.05))
    return None


def _try_alt_foreground(hwnd: int) -> None:
    win32api.keybd_event(win32con.VK_MENU, 0, 0, 0)
    try:
        win32gui.SetForegroundWindow(hwnd)
    finally:
        win32api.keybd_event(win32con.VK_MENU, 0, win32con.KEYEVENTF_KEYUP, 0)


def _attach_thread_foreground(hwnd: int) -> None:
    current_thread_id = win32api.GetCurrentThreadId()
    foreground_hwnd = win32gui.GetForegroundWindow()
    target_thread_id, _ = win32process.GetWindowThreadProcessId(hwnd)
    foreground_thread_id = 0
    if foreground_hwnd:
        foreground_thread_id, _ = win32process.GetWindowThreadProcessId(foreground_hwnd)
    attached_fg = False
    attached_target = False
    try:
        if foreground_thread_id and foreground_thread_id != current_thread_id:
            win32process.AttachThreadInput(current_thread_id, foreground_thread_id, True)
            attached_fg = True
        if target_thread_id and target_thread_id != current_thread_id:
            win32process.AttachThreadInput(current_thread_id, target_thread_id, True)
            attached_target = True
        win32gui.BringWindowToTop(hwnd)
        win32gui.SetForegroundWindow(hwnd)
    finally:
        if attached_target:
            win32process.AttachThreadInput(current_thread_id, target_thread_id, False)
        if attached_fg:
            win32process.AttachThreadInput(current_thread_id, foreground_thread_id, False)


def _force_foreground(hwnd: int, cfg: Config) -> bool:
    if win32gui.GetForegroundWindow() == hwnd:
        return True
    for attempt in (_try_alt_foreground, _attach_thread_foreground):
        try:
            attempt(hwnd)
        except Exception:
            pass
        time.sleep(cfg.kakao_foreground_retry_pause_seconds)
        if win32gui.GetForegroundWindow() == hwnd:
            return True
    return False


def activate_window(hwnd: int, cfg: Config) -> None:
    if not hwnd or not win32gui.IsWindow(hwnd):
        raise MessageSenderError("invalid window handle")
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
    else:
        win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
    time.sleep(cfg.kakao_activate_show_pause_seconds)
    ok = _force_foreground(hwnd, cfg)
    time.sleep(cfg.kakao_after_activate_pause_seconds)
    if not ok:
        raise MessageSenderError(f"failed to foreground hwnd={hwnd}")


def ensure_main_window(cfg: Config) -> int:
    if not _is_process_running(cfg.kakao_process_name):
        raise MessageSenderError(f"{cfg.kakao_process_name} is not running")
    hwnd = find_main_window(cfg)
    if hwnd:
        return hwnd
    raise MessageSenderError("process running but main window not found")


def _click_ratio(hwnd: int, ratio_x: float, ratio_y: float, pause: float) -> None:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        raise MessageSenderError("invalid window size")
    x, y = relative_point(left, top, width, height, ratio_x, ratio_y)
    pyautogui.click(x, y)
    time.sleep(pause)


def open_room(cfg: Config, room_name: str) -> int:
    main_hwnd = ensure_main_window(cfg)
    activate_window(main_hwnd, cfg)
    _click_ratio(
        main_hwnd,
        cfg.kakao_chat_tab_x,
        cfg.kakao_chat_tab_y,
        cfg.kakao_chat_tab_pause_seconds,
    )
    pyautogui.hotkey("ctrl", "f")
    time.sleep(cfg.kakao_search_open_pause_seconds)
    for _ in range(cfg.kakao_search_clear_backspace_count):
        pyautogui.press("backspace")
        time.sleep(cfg.kakao_search_reset_pause_seconds)

    pyperclip.copy(room_name)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(cfg.kakao_search_paste_pause_seconds)
    pyautogui.press("enter")
    time.sleep(cfg.kakao_room_enter_pause_seconds)

    room_hwnd = find_room_window(cfg, room_name)
    if not room_hwnd:
        room_hwnd = _wait_until(
            lambda: find_room_window(cfg, room_name),
            cfg.kakao_room_window_wait_seconds,
            cfg.kakao_window_poll_interval_seconds,
        )
    if not room_hwnd:
        raise MessageSenderError(f"room window not found: {room_name}")

    activate_window(room_hwnd, cfg)
    _click_ratio(
        room_hwnd,
        cfg.kakao_input_x,
        cfg.kakao_input_y,
        cfg.kakao_input_click_pause_seconds,
    )
    return room_hwnd


def send_text(room_name: str, text: str, cfg: Config) -> None:
    open_room(cfg, room_name)
    pyperclip.copy(text)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(cfg.kakao_paste_pause_seconds)
    pyautogui.press("enter")
    time.sleep(cfg.kakao_send_pause_seconds)
    logger.info("MESSAGE_SENT | room=%s", room_name)


def diagnose(cfg: Config) -> str:
    running = _is_process_running(cfg.kakao_process_name)
    pids = _process_pids(cfg.kakao_process_name)
    main_hwnd = find_main_window(cfg)
    room_hwnd = find_room_window(cfg, cfg.kakao_room_name)
    lines = [
        f"process_name={cfg.kakao_process_name!r}",
        f"running={running}",
        f"pids={pids}",
        f"window_class={cfg.kakao_window_class!r}",
        f"main_title={cfg.kakao_main_title!r}",
        f"room_name={cfg.kakao_room_name!r}",
    ]
    if main_hwnd:
        lines.append(
            f"main_hwnd=0x{main_hwnd:08X} title={win32gui.GetWindowText(main_hwnd)!r}"
        )
    else:
        lines.append("main_hwnd=None")
    if room_hwnd:
        lines.append(
            f"room_hwnd=0x{room_hwnd:08X} title={win32gui.GetWindowText(room_hwnd)!r}"
        )
    else:
        lines.append("room_hwnd=None")
    return "\n".join(lines)
