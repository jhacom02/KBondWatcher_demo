from __future__ import annotations

import hashlib
import logging
from typing import Optional

from pywinauto import Desktop
from pywinauto.base_wrapper import BaseWrapper

logger = logging.getLogger("kbond_watcher")


class ForestBondReaderError(RuntimeError):
    pass


def message_fingerprint(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


class ForestBondReader:
    def __init__(self, chrome_title: str) -> None:
        self.chrome_title = chrome_title
        self._watermark: set[str] = set()
        self._initialized = False

    def find_forestbond_window(self) -> BaseWrapper:
        desktop = Desktop(backend="uia")
        needle = self.chrome_title.lower()
        matches: list[BaseWrapper] = []
        try:
            windows = desktop.windows()
        except Exception as exc:
            raise ForestBondReaderError(
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
            raise ForestBondReaderError(
                f"Chrome window containing title '{self.chrome_title}' not found"
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
                raise ForestBondReaderError(
                    f"Failed to enumerate Text controls: {exc}"
                ) from exc
        return texts

    def get_visible_message_lines(self, window: Optional[BaseWrapper] = None) -> list[str]:
        win = window or self.find_forestbond_window()
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

    def initialize_watermark(
        self,
        process_existing_on_start: bool,
        lines: Optional[list[str]] = None,
    ) -> None:
        current = lines if lines is not None else self.get_visible_message_lines()
        if process_existing_on_start:
            self._watermark = set()
        else:
            self._watermark = {message_fingerprint(line) for line in current}
        self._initialized = True
        logger.info(
            "source watermark | existing=%s | process_existing=%s",
            len(current),
            process_existing_on_start,
        )

    def get_new_message_lines(
        self,
        process_existing_on_start: bool = True,
        window: Optional[BaseWrapper] = None,
    ) -> list[str]:
        lines = self.get_visible_message_lines(window=window)
        if not self._initialized:
            self.initialize_watermark(process_existing_on_start, lines=lines)
            if process_existing_on_start:
                for line in lines:
                    self._watermark.add(message_fingerprint(line))
                return lines
            return []
        new_lines: list[str] = []
        for line in lines:
            fp = message_fingerprint(line)
            if fp in self._watermark:
                continue
            self._watermark.add(fp)
            new_lines.append(line)
        return new_lines

    def diagnose(self, max_messages: int = 200) -> str:
        win = self.find_forestbond_window()
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
