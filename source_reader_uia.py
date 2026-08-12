from __future__ import annotations

import logging
from typing import Optional

from pywinauto import Desktop
from pywinauto.base_wrapper import BaseWrapper

from source_common import BaseSourceReader, SourceReaderError

logger = logging.getLogger("kbond_watcher")


class UiaSourceReader(BaseSourceReader):
    def __init__(self, source_window_title: str) -> None:
        super().__init__()
        self.source_window_title = source_window_title
        if not self.source_window_title.strip():
            raise SourceReaderError("source_window_title is required")

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
            except Exception:
                return 0

        matches.sort(key=_area, reverse=True)
        return matches[0]

    def _collect_text_controls(self, window: BaseWrapper) -> list[BaseWrapper]:
        texts: list[BaseWrapper] = []
        try:
            documents = window.descendants(control_type="Document")
        except Exception:
            documents = []
        roots: list[BaseWrapper] = list(documents) if documents else [window]
        for root in roots:
            try:
                controls = root.descendants(control_type="Text")
            except Exception as exc:
                logger.debug("Text descendants failed: %s", exc)
                controls = []
            texts.extend(controls)
        if not texts:
            try:
                texts = window.descendants(control_type="Text")
            except Exception as exc:
                raise SourceReaderError(
                    f"Failed to enumerate Text controls: {exc}"
                ) from exc
        return texts

    def get_visible_message_lines(
        self, window: Optional[BaseWrapper] = None
    ) -> list[str]:
        win = window or self.find_source_window()
        controls = self._collect_text_controls(win)
        lines: list[str] = []
        seen: set[str] = set()
        for ctrl in controls:
            try:
                raw = ctrl.window_text() or ""
            except Exception:
                continue
            for line in raw.splitlines():
                cleaned = line.strip()
                if not cleaned or cleaned in seen:
                    continue
                seen.add(cleaned)
                lines.append(cleaned)
        return lines

    def diagnose(self, max_messages: int = 200) -> str:
        win = self.find_source_window()
        try:
            title = win.window_text() or ""
        except Exception:
            title = ""
        try:
            handle = int(win.handle)
        except Exception:
            handle = 0
        controls = self._collect_text_controls(win)
        lines = self.get_visible_message_lines(window=win)
        shown = lines[: max(0, max_messages)]
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
