from __future__ import annotations

import re
import time
from typing import Optional

import win32gui
from pywinauto import Desktop
from pywinauto.base_wrapper import BaseWrapper

from .common import BaseSourceReader, SourceLine, SourceReaderError, source_line

_TIME_LINE = re.compile(
    r"^(?:(?P<sender>.+?)\s+)?\((?P<ts>\d{1,2}:\d{2}(?::\d{2})?)\)\s*[:：]?\s*$"
)
_CLOCK_TOKEN = re.compile(r"\(\d{1,2}:\d{2}")

UIA_TEXT_ENUM_MAX_RETRIES = 2
UIA_TEXT_ENUM_RETRY_DELAY_SECONDS = 0.05


def watermark_has_clock(key: str) -> bool:
    return _CLOCK_TOKEN.search(key) is not None


def attach_preceding_time(seq: list[str]) -> list[SourceLine]:
    out: list[SourceLine] = []
    prev_ts: Optional[str] = None
    for raw in seq:
        text = (raw or "").strip()
        if not text:
            continue
        time_match = _TIME_LINE.fullmatch(text)
        if time_match is not None:
            prev_ts = time_match.group("ts")
            out.append(source_line(text))
            continue
        if prev_ts is not None:
            out.append(source_line(text, f"({prev_ts}) : {text}"))
            prev_ts = None
            continue
        out.append(source_line(text))
    return out


class UiaSourceReader(BaseSourceReader):
    def __init__(self, source_window_title: str) -> None:
        super().__init__()
        self.source_window_title = source_window_title
        if not self.source_window_title.strip():
            raise SourceReaderError("source_window_title is required")
        self._window: Optional[BaseWrapper] = None

    def find_source_window(self) -> BaseWrapper:
        desktop = Desktop(backend="uia")
        needle = self.source_window_title.lower()
        matches: list[BaseWrapper] = []
        try:
            windows = desktop.windows()
        except Exception as exc:
            raise SourceReaderError(
                f"Failed to enumerate UIA desktop windows: {exc}"
            ) from exc

        for win in windows:
            try:
                title = win.window_text() or ""
            except Exception:
                continue
            if needle in title.lower():
                matches.append(win)

        if not matches:
            raise SourceReaderError(
                f"window containing title '{self.source_window_title}' not found"
            )

        def _area(w: BaseWrapper) -> int:
            try:
                rect = w.rectangle()
                return max(0, rect.width()) * max(0, rect.height())
            except Exception as exc:
                raise SourceReaderError(
                    f"Failed to read FORESTBOND window size: {exc}"
                ) from exc

        matches.sort(key=_area, reverse=True)
        self._window = matches[0]
        return self._window

    def _window_valid(self) -> bool:
        win = self._window
        if win is None:
            return False
        try:
            hwnd = int(win.handle)
        except Exception:
            return False
        return bool(hwnd and win32gui.IsWindow(hwnd) and win32gui.IsWindowVisible(hwnd))

    def _ensure_window(self) -> BaseWrapper:
        if self._window_valid() and self._window is not None:
            return self._window
        return self.find_source_window()

    def _collect_text_controls(self, window: BaseWrapper) -> list[BaseWrapper]:
        last_exc: Optional[BaseException] = None
        for attempt in range(UIA_TEXT_ENUM_MAX_RETRIES + 1):
            try:
                texts = window.descendants(control_type="Text")
            except Exception as exc:
                last_exc = exc
                if attempt < UIA_TEXT_ENUM_MAX_RETRIES:
                    time.sleep(UIA_TEXT_ENUM_RETRY_DELAY_SECONDS)
                    continue
                raise SourceReaderError(
                    f"Failed to enumerate Text controls: {exc}"
                ) from exc
            if not texts:
                raise SourceReaderError("no UIA Text controls")
            return texts
        raise SourceReaderError(
            f"Failed to enumerate Text controls: {last_exc}"
        )

    def get_visible_message_lines(
        self, window: Optional[BaseWrapper] = None
    ) -> list[SourceLine]:
        win = window or self._ensure_window()
        controls = self._collect_text_controls(win)
        raw_lines: list[str] = []
        for ctrl in controls:
            try:
                raw = ctrl.window_text() or ""
            except Exception as exc:
                raise SourceReaderError(f"Failed to read UIA Text: {exc}") from exc
            for line in raw.splitlines():
                cleaned = line.strip()
                if cleaned:
                    raw_lines.append(cleaned)
        paired = attach_preceding_time(raw_lines)
        lines: list[SourceLine] = []
        seen: set[str] = set()
        for item in paired:
            if item.watermark_key in seen:
                continue
            seen.add(item.watermark_key)
            lines.append(item)
        return lines

    def diagnose(self, max_messages: int = 200) -> str:
        win = self.find_source_window()
        title = win.window_text() or ""
        handle = int(win.handle)
        controls = self._collect_text_controls(win)
        lines = self.get_visible_message_lines(window=win)
        shown = [item.watermark_key for item in lines[: max(0, max_messages)]]
        return "\n".join(
            [
                f"window title: {title!r}",
                f"HWND: 0x{handle:08X} ({handle})",
                f"Text control count: {len(controls)}",
                f"message lines: {len(lines)} (showing {len(shown)})",
                "---- messages ----",
                *shown,
            ]
        )
