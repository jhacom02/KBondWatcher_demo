from __future__ import annotations

import hashlib
import logging
from typing import Optional

import psutil
import win32gui
import win32process
from pywinauto import Application, Desktop
from pywinauto.base_wrapper import BaseWrapper

logger = logging.getLogger("kbond_watcher")


class SourceReaderError(RuntimeError):
    pass


def message_fingerprint(text: str) -> str:
    return hashlib.sha1(text.encode("utf-8")).hexdigest()


def _process_pids(process_name: str) -> set[int]:
    expected = process_name.lower()
    pids: set[int] = set()
    for proc in psutil.process_iter(["name", "pid"]):
        try:
            if (proc.info.get("name") or "").lower() == expected:
                pids.add(int(proc.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids


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


class SourceReader:
    def __init__(
        self,
        source_window_title: str,
        source_process_name: str = "",
    ) -> None:
        self.source_window_title = source_window_title
        self.source_process_name = source_process_name.strip()
        self._watermark: set[str] = set()
        self._initialized = False
        self._hwnd: Optional[int] = None

    def _find_hwnd(self) -> Optional[int]:
        needle = self.source_window_title.lower()
        pids: Optional[set[int]] = None
        if self.source_process_name:
            pids = _process_pids(self.source_process_name)
            if not pids:
                return None

        matches: list[int] = []

        def _callback(hwnd: int, _: object) -> bool:
            try:
                title = win32gui.GetWindowText(hwnd) or ""
                if needle not in title.lower():
                    return True
                if pids is not None:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    if int(pid) not in pids:
                        return True
                matches.append(hwnd)
            except Exception:
                pass
            return True

        win32gui.EnumWindows(_callback, None)
        if not matches:
            return None
        return max(matches, key=_rank_window)

    def _wrap_hwnd(self, hwnd: int) -> BaseWrapper:
        try:
            app = Application(backend="uia").connect(handle=hwnd)
            return app.window(handle=hwnd)
        except Exception:
            desktop = Desktop(backend="uia")
            for win in desktop.windows():
                try:
                    if int(win.handle) == int(hwnd):
                        return win
                except Exception:
                    continue
            raise SourceReaderError(f"Failed to wrap HWND 0x{hwnd:08X} with UIA")

    def find_source_window(self) -> BaseWrapper:
        hwnd = self._find_hwnd()
        if hwnd is not None:
            self._hwnd = hwnd
            return self._wrap_hwnd(hwnd)

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
            detail = f"window containing title '{self.source_window_title}' not found"
            if self.source_process_name:
                detail += f" (process={self.source_process_name!r})"
            raise SourceReaderError(detail)

        def _area(w: BaseWrapper) -> int:
            try:
                rect = w.rectangle()
                return max(0, rect.width()) * max(0, rect.height())
            except Exception:
                return 0

        matches.sort(key=_area, reverse=True)
        try:
            self._hwnd = int(matches[0].handle)
        except Exception:
            self._hwnd = None
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
                logger.debug("Text descendants failed: %s", exc)
                texts = []
        if not texts:
            for control_type in ("ListItem", "Edit", "DataItem"):
                try:
                    extra = window.descendants(control_type=control_type)
                except Exception as exc:
                    logger.debug("%s descendants failed: %s", control_type, exc)
                    continue
                texts.extend(extra)
        return texts

    def _collect_win32_lines(self, hwnd: int) -> list[str]:
        skip_classes = {
            "tedit",
            "edit",
            "tbuttonededit",
            "ttntcombobox",
            "combobox",
            "tezflatcheckboxw",
            "button",
            "tbutton",
        }
        skip_exact = {
            "pnlmain",
            "pnlmain_class",
            "flatpanel1",
            "panel5",
            "panel8",
            "pnladdroom",
        }
        lines: list[str] = []
        seen: set[str] = set()
        try:
            app = Application(backend="win32").connect(handle=hwnd)
            win = app.window(handle=hwnd)
            controls = win.descendants()
        except Exception as exc:
            logger.debug("win32 descendants failed: %s", exc)
            return lines

        for ctrl in controls:
            try:
                cls = (ctrl.class_name() or "").lower()
            except Exception:
                cls = ""
            if cls in skip_classes:
                continue
            try:
                if not ctrl.is_visible():
                    continue
            except Exception:
                pass
            chunks: list[str] = []
            try:
                texts = ctrl.texts()
                if texts:
                    chunks.extend(str(t) for t in texts if t)
            except Exception:
                pass
            try:
                raw = ctrl.window_text() or ""
                if raw:
                    chunks.append(raw)
            except Exception:
                pass
            for chunk in chunks:
                for line in str(chunk).splitlines():
                    cleaned = line.strip()
                    if not cleaned or cleaned in seen:
                        continue
                    if cleaned.lower() in skip_exact:
                        continue
                    if cleaned.startswith("pnl") and "_" in cleaned.lower():
                        continue
                    seen.add(cleaned)
                    lines.append(cleaned)
        return lines

    def get_visible_message_lines(self, window: Optional[BaseWrapper] = None) -> list[str]:
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

        if not lines:
            hwnd = self._hwnd
            if hwnd is None:
                try:
                    hwnd = int(win.handle)
                except Exception:
                    hwnd = None
            if hwnd is not None:
                for line in self._collect_win32_lines(hwnd):
                    if line not in seen:
                        seen.add(line)
                        lines.append(line)
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
        win = self.find_source_window()
        try:
            title = win.window_text() or ""
        except Exception:
            title = ""
        try:
            handle = int(win.handle)
        except Exception:
            handle = self._hwnd or 0
        controls = self._collect_text_controls(win)
        lines = self.get_visible_message_lines(window=win)
        shown = lines[: max(0, max_messages)]
        return "\n".join(
            [
                f"window title: {title!r}",
                f"HWND: 0x{handle:08X} ({handle})",
                f"process: {self.source_process_name or '(any)'}",
                f"Text control count: {len(controls)}",
                f"message lines: {len(lines)} (showing {len(shown)})",
                "---- messages ----",
                *shown,
            ]
        )
