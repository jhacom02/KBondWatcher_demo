from __future__ import annotations

from typing import Optional

import win32gui

from .richedit import (
    CHAT_EDIT_CLASS,
    RichEditReaderError,
    find_chat_richedit,
    is_gettext_clipped,
    read_richedit_length,
    read_richedit_snapshot,
    resolve_gettext_read,
)
from .common import BaseSourceReader, SourceLine, SourceReaderError, as_source_lines


class KbondSourceReader(BaseSourceReader):
    def __init__(
        self,
        source_window_title: str,
        source_process_name: str,
    ) -> None:
        super().__init__()
        self.source_window_title = source_window_title
        self.source_process_name = source_process_name.strip()
        if not self.source_process_name:
            raise SourceReaderError("source_process_name is required")
        if not self.source_window_title.strip():
            raise SourceReaderError("source_window_title is required")
        self._hwnd: Optional[int] = None
        self._edit_hwnd: Optional[int] = None
        self._cache_len: Optional[int] = None
        self._cached_lines: list[str] = []
        self._last_gettext_len: int = -1
        self._last_clipped: bool = False

    def find_source_window(self) -> int:
        try:
            parent, edit = find_chat_richedit(
                self.source_process_name,
                self.source_window_title,
            )
        except RichEditReaderError as exc:
            raise SourceReaderError(str(exc)) from exc
        if edit != self._edit_hwnd:
            self._invalidate_richedit_cache()
        self._hwnd = parent
        self._edit_hwnd = edit
        return parent

    def _invalidate_richedit_cache(self) -> None:
        self._cache_len = None
        self._cached_lines = []
        self._last_gettext_len = -1
        self._last_clipped = False

    def _chat_valid(self) -> bool:
        return bool(
            self._edit_hwnd
            and win32gui.IsWindow(self._edit_hwnd)
            and win32gui.IsWindowVisible(self._edit_hwnd)
        )

    def _ensure_chat(self) -> None:
        if self._chat_valid():
            return
        self.find_source_window()

    def _read_richedit_lines(self) -> list[str]:
        if self._edit_hwnd is None:
            raise SourceReaderError("TJvRichEdit hwnd not resolved")
        try:
            char_len = read_richedit_length(self._edit_hwnd)
        except RichEditReaderError as exc:
            raise SourceReaderError(str(exc)) from exc
        self._last_gettext_len = char_len
        action = resolve_gettext_read(char_len, self._cache_len)
        if action == "use_cache":
            return list(self._cached_lines)
        try:
            snap = read_richedit_snapshot(self._edit_hwnd)
        except RichEditReaderError as exc:
            raise SourceReaderError(str(exc)) from exc
        self._last_gettext_len = snap.char_len
        self._last_clipped = snap.clipped
        self._cache_len = snap.char_len
        self._cached_lines = snap.lines
        return list(self._cached_lines)

    def get_visible_message_lines(self) -> list[SourceLine]:
        self._ensure_chat()
        return as_source_lines(self._read_richedit_lines())

    def diagnose(self, max_messages: int = 200) -> str:
        hwnd = self.find_source_window()
        title = win32gui.GetWindowText(hwnd) or ""
        parent_class = win32gui.GetClassName(hwnd) or ""
        ctrl = self._edit_hwnd or 0
        ctrl_class = ""
        ctrl_rect = ""
        if ctrl:
            ctrl_class = win32gui.GetClassName(ctrl) or ""
            left, top, right, bottom = win32gui.GetWindowRect(ctrl)
            ctrl_rect = f"({left}, {top}, {right}, {bottom})"
        lines = self.get_visible_message_lines()
        shown = [item.watermark_key for item in lines[: max(0, max_messages)]]
        out = [
            f"backend: richedit",
            f"chat title needle: {self.source_window_title!r}",
            f"window title: {title!r}",
            f"window class: {parent_class!r}",
            f"HWND: 0x{hwnd:08X} ({hwnd})",
            f"process: {self.source_process_name}",
            f"chat class: {(ctrl_class or CHAT_EDIT_CLASS)!r}",
            f"chat HWND: 0x{ctrl:08X} ({ctrl})",
            f"chat rect: {ctrl_rect or '(n/a)'}",
        ]
        if ctrl:
            gettext_len = read_richedit_length(ctrl)
            out.append(f"gettext_len: {gettext_len}")
            out.append(f"clipped: {is_gettext_clipped(gettext_len)}")
        out.append(f"message lines: {len(lines)} (showing {len(shown)})")
        out.append("---- messages ----")
        out.extend(shown)
        return "\n".join(out)
