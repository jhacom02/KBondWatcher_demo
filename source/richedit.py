from __future__ import annotations

import ctypes
from ctypes import wintypes
from dataclasses import dataclass
from typing import Optional

import win32con
import win32gui
import win32process

from .eltree import normalize_lines, process_pids

CHAT_EDIT_CLASS = "TJvRichEdit"
MIN_CHAT_AREA = 20_000
MIN_CHAT_HEIGHT = 80
MAX_TEXT_CHARS = 2_000_000

WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E

user32 = ctypes.WinDLL("user32", use_last_error=True)
user32.SendMessageW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]
user32.SendMessageW.restype = wintypes.LPARAM


class RichEditReaderError(RuntimeError):
    pass


@dataclass(frozen=True)
class RichEditSnapshot:
    lines: list[str]
    char_len: int
    clipped: bool


def is_gettext_clipped(char_len: int, max_chars: int = MAX_TEXT_CHARS) -> bool:
    return char_len > max_chars


def resolve_gettext_read(
    char_len: int,
    cached_len: Optional[int],
    max_chars: int = MAX_TEXT_CHARS,
) -> str:
    """Return 'skip_clip', 'use_cache', or 'fetch'."""
    if is_gettext_clipped(char_len, max_chars):
        return "skip_clip"
    if cached_len is not None and char_len == cached_len:
        return "use_cache"
    return "fetch"


def read_richedit_length(hwnd: int) -> int:
    return int(user32.SendMessageW(hwnd, WM_GETTEXTLENGTH, 0, 0))


@dataclass(frozen=True)
class EditCandidate:
    hwnd: int
    parent_hwnd: int
    parent_title: str
    parent_class: str
    left: int
    top: int
    right: int
    bottom: int
    visible: bool = True

    @property
    def area(self) -> int:
        return max(0, self.right - self.left) * max(0, self.bottom - self.top)

    @property
    def height(self) -> int:
        return max(0, self.bottom - self.top)


def pick_chat_edit(
    candidates: list[EditCandidate],
    min_area: int = MIN_CHAT_AREA,
    min_height: int = MIN_CHAT_HEIGHT,
) -> Optional[EditCandidate]:
    eligible = [c for c in candidates if c.area >= min_area and c.height >= min_height]
    if not eligible:
        return None
    return max(eligible, key=lambda c: c.area)


def _root_hwnd(hwnd: int) -> int:
    try:
        return int(win32gui.GetAncestor(hwnd, win32con.GA_ROOT))
    except Exception:
        return int(win32gui.GetParent(hwnd) or hwnd)


def list_richedit_candidates(process_name: str) -> list[EditCandidate]:
    pids = process_pids(process_name)
    if not pids:
        raise RichEditReaderError(f"process not running: {process_name!r}")

    found: list[EditCandidate] = []
    parents: list[int] = []

    def _top(hwnd: int, _: object) -> bool:
        try:
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if int(pid) in pids:
                parents.append(int(hwnd))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_top, None)

    def _child(hwnd: int, _: object) -> bool:
        try:
            cls = win32gui.GetClassName(hwnd) or ""
            if not cls.startswith(CHAT_EDIT_CLASS):
                return True
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            root = _root_hwnd(hwnd)
            found.append(
                EditCandidate(
                    hwnd=int(hwnd),
                    parent_hwnd=root,
                    parent_title=win32gui.GetWindowText(root) or "",
                    parent_class=win32gui.GetClassName(root) or "",
                    left=left,
                    top=top,
                    right=right,
                    bottom=bottom,
                    visible=bool(win32gui.IsWindowVisible(hwnd)),
                )
            )
        except Exception:
            pass
        return True

    for parent in parents:
        win32gui.EnumChildWindows(parent, _child, None)
    return found


def find_chat_richedit(process_name: str) -> tuple[int, int]:
    candidates = list_richedit_candidates(process_name)
    visible = [c for c in candidates if c.visible]
    chosen = pick_chat_edit(visible) or pick_chat_edit(candidates)
    if chosen is None:
        raise RichEditReaderError(
            f"no visible {CHAT_EDIT_CLASS} chat pane "
            f"(visible={len(visible)} total={len(candidates)})"
        )
    return chosen.parent_hwnd, chosen.hwnd


def read_richedit_text(hwnd: int, max_chars: int = MAX_TEXT_CHARS) -> tuple[str, int, bool]:
    try:
        n = read_richedit_length(hwnd)
    except Exception as exc:
        raise RichEditReaderError(f"WM_GETTEXTLENGTH failed: {exc}") from exc
    if n <= 0:
        return "", 0, False
    if is_gettext_clipped(n, max_chars):
        return "", n, True
    buf = ctypes.create_unicode_buffer(n + 1)
    try:
        user32.SendMessageW(hwnd, WM_GETTEXT, n + 1, ctypes.addressof(buf))
    except Exception as exc:
        raise RichEditReaderError(f"WM_GETTEXT failed: {exc}") from exc
    return buf.value or "", n, False


def read_richedit_snapshot(
    hwnd: int, max_chars: int = MAX_TEXT_CHARS
) -> RichEditSnapshot:
    text, char_len, clipped = read_richedit_text(hwnd, max_chars=max_chars)
    if clipped:
        return RichEditSnapshot(lines=[], char_len=char_len, clipped=True)
    return RichEditSnapshot(
        lines=normalize_lines([text]),
        char_len=char_len,
        clipped=False,
    )


def read_richedit_lines(hwnd: int) -> list[str]:
    snap = read_richedit_snapshot(hwnd)
    return snap.lines
