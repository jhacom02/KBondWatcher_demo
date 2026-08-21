from __future__ import annotations

import time
from collections import Counter
from typing import Optional

import win32gui
from pywinauto import Desktop
from pywinauto.base_wrapper import BaseWrapper

from .common import (
    BaseSourceReader,
    SourceLine,
    SourceReaderError,
    source_line,
    watermark_window,
)

UIA_TEXT_ENUM_MAX_RETRIES = 2
UIA_TEXT_ENUM_RETRY_DELAY_SECONDS = 0.05


def new_lines_after(
    prev: list[SourceLine],
    now: list[SourceLine],
    max_counts: Optional[Counter[str]] = None,
) -> list[SourceLine]:
    if not prev:
        return []
    prev_keys = [item.watermark_key for item in prev]
    now_keys = [item.watermark_key for item in now]
    if not (set(prev_keys) & set(now_keys)):
        return []
    baseline = max_counts if max_counts is not None else Counter(prev_keys)
    need: dict[str, int] = {}
    for key, count in Counter(now_keys).items():
        extra = count - baseline[key]
        if extra > 0:
            need[key] = extra
    new: list[SourceLine] = []
    for item in now:
        key = item.watermark_key
        left = need.get(key, 0)
        if left <= 0:
            continue
        new.append(item)
        need[key] = left - 1
    return new


def bump_max_counts(max_counts: Counter[str], lines: list[SourceLine]) -> None:
    for key, count in Counter(item.watermark_key for item in lines).items():
        if count > max_counts[key]:
            max_counts[key] = count


class UiaSourceReader(BaseSourceReader):
    def __init__(self, source_window_title: str) -> None:
        super().__init__()
        self.source_window_title = source_window_title
        if not self.source_window_title.strip():
            raise SourceReaderError("source_window_title is required")
        self._window: Optional[BaseWrapper] = None
        self._prev_snapshot: list[SourceLine] = []
        self._max_counts: Counter[str] = Counter()

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
        return [source_line(raw) for raw in raw_lines]

    def initialize_watermark(
        self,
        process_existing_on_start: bool,
        lines: Optional[list[SourceLine]] = None,
    ) -> None:
        current = lines if lines is not None else self.get_visible_message_lines()
        self._prev_snapshot = watermark_window(current)
        self._max_counts = Counter(
            item.watermark_key for item in self._prev_snapshot
        )
        self._initialized = True

    def reseed_watermark_from_visible(self) -> None:
        return

    def get_new_message_lines(
        self,
        process_existing_on_start: bool = False,
    ) -> list[SourceLine]:
        now = watermark_window(self.get_visible_message_lines())
        if not self._initialized:
            self.initialize_watermark(process_existing_on_start, lines=now)
            return []
        new = new_lines_after(self._prev_snapshot, now, self._max_counts)
        bump_max_counts(self._max_counts, now)
        self._prev_snapshot = now
        return new

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
