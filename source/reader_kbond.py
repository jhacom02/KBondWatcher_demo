from __future__ import annotations

import logging
from typing import Optional

import win32gui

from .eltree import (
    TREE_CLASS,
    ElTreeReaderError,
    find_chat_eltree,
    find_messenger_hwnd,
    item_count,
    read_eltree_lines,
)
from .richedit import (
    CHAT_EDIT_CLASS,
    MAX_TEXT_CHARS,
    RichEditReaderError,
    find_chat_richedit,
    is_gettext_clipped,
    read_richedit_length,
    read_richedit_snapshot,
    resolve_gettext_read,
)
from .common import BaseSourceReader, SourceLine, SourceReaderError, as_source_lines

logger = logging.getLogger("kbond_watcher")


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
        self._tree_hwnd: Optional[int] = None
        self._backend: Optional[str] = None
        self._cache_len: Optional[int] = None
        self._cached_lines: list[str] = []
        self._last_gettext_len: int = -1
        self._last_clipped: bool = False
        self._clip_logged: bool = False

    def find_source_window(self) -> int:
        richedit_error: Optional[str] = None
        try:
            parent, edit = find_chat_richedit(self.source_process_name)
            if edit != self._edit_hwnd:
                self._invalidate_richedit_cache()
            self._hwnd = parent
            self._edit_hwnd = edit
            self._tree_hwnd = None
            self._backend = "richedit"
            return parent
        except RichEditReaderError as exc:
            richedit_error = str(exc)

        try:
            hwnd = find_messenger_hwnd(
                self.source_process_name,
                self.source_window_title,
            )
            tree = find_chat_eltree(hwnd)
        except ElTreeReaderError as exc:
            detail = str(exc)
            if richedit_error:
                detail = f"{richedit_error}; {detail}"
            raise SourceReaderError(detail) from exc
        self._hwnd = hwnd
        self._edit_hwnd = None
        self._tree_hwnd = tree
        self._backend = "eltree"
        self._invalidate_richedit_cache()
        return hwnd

    def _invalidate_richedit_cache(self) -> None:
        self._cache_len = None
        self._cached_lines = []
        self._last_gettext_len = -1
        self._last_clipped = False
        self._clip_logged = False

    def _chat_valid(self) -> bool:
        if self._backend == "richedit":
            return bool(
                self._edit_hwnd
                and win32gui.IsWindow(self._edit_hwnd)
                and win32gui.IsWindowVisible(self._edit_hwnd)
            )
        if self._backend == "eltree":
            return bool(self._tree_hwnd and win32gui.IsWindow(self._tree_hwnd))
        return False

    def _ensure_chat(self) -> None:
        if self._chat_valid():
            return
        self.find_source_window()

    def _read_richedit_lines(self) -> list[str]:
        if self._edit_hwnd is None:
            raise SourceReaderError("TJvRichEdit hwnd not resolved")
        char_len = read_richedit_length(self._edit_hwnd)
        self._last_gettext_len = char_len
        action = resolve_gettext_read(char_len, self._cache_len)
        if action == "skip_clip":
            self._last_clipped = True
            if not self._clip_logged:
                logger.warning(
                    "WM_GETTEXT clipped | len=%s cap=%s — skip truncated head",
                    char_len,
                    MAX_TEXT_CHARS,
                )
                self._clip_logged = True
            return list(self._cached_lines)
        self._last_clipped = False
        self._clip_logged = False
        if action == "use_cache":
            return list(self._cached_lines)
        snap = read_richedit_snapshot(self._edit_hwnd)
        self._last_gettext_len = snap.char_len
        if snap.clipped:
            self._last_clipped = True
            if not self._clip_logged:
                logger.warning(
                    "WM_GETTEXT clipped | len=%s cap=%s — skip truncated head",
                    snap.char_len,
                    MAX_TEXT_CHARS,
                )
                self._clip_logged = True
            return list(self._cached_lines)
        self._cache_len = snap.char_len
        self._cached_lines = snap.lines
        self._last_clipped = False
        return list(self._cached_lines)

    def get_visible_message_lines(self) -> list[SourceLine]:
        self._ensure_chat()
        try:
            if self._backend == "richedit":
                return as_source_lines(self._read_richedit_lines())
            if self._tree_hwnd is None:
                raise SourceReaderError("TElTree hwnd not resolved")
            return as_source_lines(read_eltree_lines(self._tree_hwnd))
        except (RichEditReaderError, ElTreeReaderError) as exc:
            raise SourceReaderError(str(exc)) from exc

    def diagnose(self, max_messages: int = 200) -> str:
        hwnd = self.find_source_window()
        title = win32gui.GetWindowText(hwnd) or ""
        parent_class = win32gui.GetClassName(hwnd) or ""
        ctrl = self._edit_hwnd or self._tree_hwnd or 0
        ctrl_class = ""
        ctrl_rect = ""
        count = -1
        if ctrl:
            ctrl_class = win32gui.GetClassName(ctrl) or ""
            left, top, right, bottom = win32gui.GetWindowRect(ctrl)
            ctrl_rect = f"({left}, {top}, {right}, {bottom})"
            if self._backend == "eltree":
                count = item_count(ctrl)
        try:
            lines = self.get_visible_message_lines()
            read_error = ""
        except SourceReaderError as exc:
            lines = []
            read_error = str(exc)
        shown = [item.watermark_key for item in lines[: max(0, max_messages)]]
        default_class = CHAT_EDIT_CLASS if self._backend == "richedit" else TREE_CLASS
        out = [
            f"backend: {self._backend or 'unknown'}",
            f"window title: {title!r}",
            f"window class: {parent_class!r}",
            f"HWND: 0x{hwnd:08X} ({hwnd})",
            f"process: {self.source_process_name}",
            f"chat class: {(ctrl_class or default_class)!r}",
            f"chat HWND: 0x{ctrl:08X} ({ctrl})",
            f"chat rect: {ctrl_rect or '(n/a)'}",
        ]
        if self._backend == "eltree":
            out.append(f"item count (TVM_GETCOUNT): {count}")
        if self._backend == "richedit" and ctrl:
            gettext_len = self._last_gettext_len
            clipped = self._last_clipped
            try:
                gettext_len = read_richedit_length(ctrl)
                clipped = is_gettext_clipped(gettext_len)
            except Exception:
                pass
            out.append(f"gettext_len: {gettext_len}")
            out.append(f"clipped: {clipped}")
        out.append(f"message lines: {len(lines)} (showing {len(shown)})")
        if read_error:
            out.append(f"read error: {read_error}")
        out.append("---- messages ----")
        out.extend(shown)
        return "\n".join(out)
