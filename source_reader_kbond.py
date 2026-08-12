from __future__ import annotations

from typing import Optional

import win32gui

from eltree_reader import (
    TREE_CLASS,
    ElTreeReaderError,
    find_chat_eltree,
    find_messenger_hwnd,
    item_count,
    read_eltree_lines,
)
from source_common import BaseSourceReader, SourceReaderError


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
        self._tree_hwnd: Optional[int] = None

    def find_source_window(self) -> int:
        try:
            hwnd = find_messenger_hwnd(
                self.source_process_name,
                self.source_window_title,
            )
            tree = find_chat_eltree(hwnd)
        except ElTreeReaderError as exc:
            raise SourceReaderError(str(exc)) from exc
        self._hwnd = hwnd
        self._tree_hwnd = tree
        return hwnd

    def _ensure_tree(self) -> int:
        if self._tree_hwnd and win32gui.IsWindow(self._tree_hwnd):
            return self._tree_hwnd
        self.find_source_window()
        if self._tree_hwnd is None:
            raise SourceReaderError("TElTree hwnd not resolved")
        return self._tree_hwnd

    def get_visible_message_lines(self) -> list[str]:
        tree = self._ensure_tree()
        try:
            return read_eltree_lines(tree)
        except ElTreeReaderError as exc:
            raise SourceReaderError(str(exc)) from exc

    def diagnose(self, max_messages: int = 200) -> str:
        hwnd = self.find_source_window()
        tree = self._tree_hwnd or 0
        title = win32gui.GetWindowText(hwnd) or ""
        tree_class = ""
        tree_rect = ""
        count = -1
        if tree:
            tree_class = win32gui.GetClassName(tree) or ""
            left, top, right, bottom = win32gui.GetWindowRect(tree)
            tree_rect = f"({left}, {top}, {right}, {bottom})"
            count = item_count(tree)
        try:
            lines = self.get_visible_message_lines()
            read_error = ""
        except SourceReaderError as exc:
            lines = []
            read_error = str(exc)
        shown = lines[: max(0, max_messages)]
        out = [
            f"window title: {title!r}",
            f"HWND: 0x{hwnd:08X} ({hwnd})",
            f"process: {self.source_process_name}",
            f"tree class: {(tree_class or TREE_CLASS)!r}",
            f"tree HWND: 0x{tree:08X} ({tree})",
            f"tree rect: {tree_rect or '(n/a)'}",
            f"item count (TVM_GETCOUNT): {count}",
            f"message lines: {len(lines)} (showing {len(shown)})",
        ]
        if read_error:
            out.append(f"read error: {read_error}")
        out.append("---- messages ----")
        out.extend(shown)
        return "\n".join(out)
