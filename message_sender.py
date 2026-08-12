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


def _is_target_window(hwnd: int, cfg: Config, pids: set[int]) -> bool:
    title = win32gui.GetWindowText(hwnd) or ""
    if cfg.send_window_title.lower() not in title.lower():
        return False
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    return int(pid) in pids


def find_target_window(cfg: Config) -> Optional[int]:
    pids = set(_process_pids(cfg.send_process_name))
    if not pids:
        return None
    return _find_best(lambda hwnd: _is_target_window(hwnd, cfg, pids))


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
        time.sleep(cfg.send_foreground_retry_pause_seconds)
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
    time.sleep(cfg.send_activate_show_pause_seconds)
    ok = _force_foreground(hwnd, cfg)
    time.sleep(cfg.send_after_activate_pause_seconds)
    if not ok:
        raise MessageSenderError(f"failed to foreground hwnd={hwnd}")


def _click_ratio(hwnd: int, ratio_x: float, ratio_y: float, pause: float) -> None:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        raise MessageSenderError("invalid window size")
    x, y = relative_point(left, top, width, height, ratio_x, ratio_y)
    pyautogui.click(x, y)
    time.sleep(pause)


def ensure_target_window(cfg: Config) -> int:
    if not _is_process_running(cfg.send_process_name):
        raise MessageSenderError(f"{cfg.send_process_name} is not running")
    hwnd = find_target_window(cfg)
    if hwnd:
        return hwnd
    raise MessageSenderError(
        f"window containing title {cfg.send_window_title!r} not found"
    )


def send_text(text: str, cfg: Config) -> None:
    hwnd = ensure_target_window(cfg)
    activate_window(hwnd, cfg)
    _click_ratio(
        hwnd,
        cfg.send_input_x,
        cfg.send_input_y,
        cfg.send_input_click_pause_seconds,
    )
    pyperclip.copy(text)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(cfg.send_paste_pause_seconds)
    pyautogui.press("enter")
    time.sleep(cfg.send_send_pause_seconds)
    logger.info("MESSAGE_SENT | title=%s", cfg.send_window_title)


def diagnose(cfg: Config) -> str:
    running = _is_process_running(cfg.send_process_name)
    pids = _process_pids(cfg.send_process_name)
    hwnd = find_target_window(cfg)
    lines = [
        f"process_name={cfg.send_process_name!r}",
        f"running={running}",
        f"pids={pids}",
        f"window_title_needle={cfg.send_window_title!r}",
        f"input_ratio=({cfg.send_input_x}, {cfg.send_input_y})",
    ]
    if hwnd:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top
        click_x, click_y = relative_point(
            left, top, width, height, cfg.send_input_x, cfg.send_input_y
        )
        lines.append(f"hwnd=0x{hwnd:08X} title={win32gui.GetWindowText(hwnd)!r}")
        lines.append(f"rect=({left}, {top}, {right}, {bottom})")
        lines.append(f"click_point=({click_x}, {click_y})")
    else:
        lines.append("hwnd=None")
    return "\n".join(lines)
