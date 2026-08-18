from __future__ import annotations

import time
from typing import Optional

import win32api
import win32gui

from send.ui import find_target_window
from config import Config


class CalibrationError(RuntimeError):
    pass


def capture_click_ratio(cfg: Config, timeout_seconds: float = 20.0) -> tuple[float, float]:
    """Wait for left-click and return ratios relative to send target window."""
    hwnd = find_target_window(cfg)
    if not hwnd:
        raise CalibrationError("send target window not found")
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        raise CalibrationError("invalid send window size")

    deadline = time.monotonic() + float(timeout_seconds)
    was_down = False
    while time.monotonic() < deadline:
        down = bool(win32api.GetAsyncKeyState(0x01) & 0x8000)
        if down and not was_down:
            x, y = win32api.GetCursorPos()
            if not (left <= x <= right and top <= y <= bottom):
                was_down = down
                time.sleep(0.05)
                continue
            # wait release
            while win32api.GetAsyncKeyState(0x01) & 0x8000:
                time.sleep(0.01)
            rx = max(0.0, min(1.0, (x - left) / float(width)))
            ry = max(0.0, min(1.0, (y - top) / float(height)))
            return rx, ry
        was_down = down
        time.sleep(0.02)
    raise CalibrationError("timed out waiting for click")
