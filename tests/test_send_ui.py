from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from send.ui import activate_window, relative_point, send_text  # noqa: E402


def test_relative_point_center() -> None:
    assert relative_point(0, 0, 1000, 800, 0.5, 0.5) == (500, 400)


def test_relative_point_input() -> None:
    assert relative_point(100, 200, 400, 600, 0.5, 0.5) == (300, 500)


def test_activate_window_skips_sleep_when_already_foreground(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    sleeps: list[float] = []
    monkeypatch.setattr("send.ui.time.sleep", lambda s: sleeps.append(s))
    monkeypatch.setattr("send.ui.win32gui.IsWindow", lambda hwnd: True)
    monkeypatch.setattr("send.ui.win32gui.IsIconic", lambda hwnd: False)
    monkeypatch.setattr("send.ui.win32gui.GetForegroundWindow", lambda: 1)
    monkeypatch.setattr("send.ui.win32gui.ShowWindow", lambda *a: None)
    monkeypatch.setattr("send.ui._force_foreground", lambda hwnd, cfg: True)

    class _Cfg:
        send_activate_show_pause_seconds = 0.05
        send_after_activate_pause_seconds = 0.10

    activate_window(1, _Cfg())
    assert sleeps == []


def test_send_text_does_not_reforce_after_activate(monkeypatch: pytest.MonkeyPatch) -> None:
    calls = {"force": 0, "activate": 0}

    def _activate(hwnd: int, cfg: object) -> None:
        calls["activate"] += 1

    def _force(hwnd: int, cfg: object) -> bool:
        calls["force"] += 1
        return True

    class _Cfg:
        send_process_name = "notepad.exe"
        send_window_title = "메모장"
        send_input_x = 0.5
        send_input_y = 0.5
        send_input_click_pause_seconds = 0
        send_paste_pause_seconds = 0
        send_send_pause_seconds = 0

    monkeypatch.setattr("send.ui.ensure_target_window", lambda cfg: 1)
    monkeypatch.setattr("send.ui.activate_window", _activate)
    monkeypatch.setattr("send.ui._force_foreground", _force)
    monkeypatch.setattr("send.ui._set_topmost", lambda hwnd, enabled: None)
    monkeypatch.setattr("send.ui._click_ratio", lambda *a, **k: None)
    monkeypatch.setattr("send.ui._same_app", lambda a, b: True)
    monkeypatch.setattr("send.ui.win32gui.GetForegroundWindow", lambda: 1)
    monkeypatch.setattr("send.ui.win32gui.WindowFromPoint", lambda pos: 1)
    monkeypatch.setattr("send.ui.win32gui.GetWindowText", lambda hwnd: "메모장")
    monkeypatch.setattr("send.ui.win32api.GetCursorPos", lambda: (0, 0))
    monkeypatch.setattr("send.ui.pyperclip.copy", lambda text: None)
    monkeypatch.setattr("send.ui.pyautogui.hotkey", lambda *a: None)
    monkeypatch.setattr("send.ui.pyautogui.press", lambda *a: None)
    send_text("hi", _Cfg())
    assert calls["activate"] == 1
    assert calls["force"] == 0
