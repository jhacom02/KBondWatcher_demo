from __future__ import annotations

import ctypes
import struct
from ctypes import wintypes
from dataclasses import dataclass
from typing import Optional

import win32con
import win32gui
import win32process

from .win32mem import (
    MEM_COMMIT,
    MEM_RELEASE,
    MEM_RESERVE,
    PAGE_READWRITE,
    PROCESS_ACCESS,
    kernel32,
    normalize_lines,
    process_is_32bit,
    process_pids,
)

CHAT_EDIT_CLASS = "TJvRichEdit"
MIN_CHAT_AREA = 20_000
MIN_CHAT_HEIGHT = 80
MAX_TEXT_CHARS = 200_000

WM_GETTEXT = 0x000D
WM_GETTEXTLENGTH = 0x000E
WM_USER = 0x0400
EM_GETTEXTRANGE = WM_USER + 75

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


def title_matches(title: str, needle: str) -> bool:
    text = (needle or "").strip()
    if not text:
        return False
    return text.lower() in (title or "").lower()


def pick_chat_edit(
    candidates: list[EditCandidate],
    min_area: int = MIN_CHAT_AREA,
    min_height: int = MIN_CHAT_HEIGHT,
) -> Optional[EditCandidate]:
    eligible = [c for c in candidates if c.area >= min_area and c.height >= min_height]
    if not eligible:
        return None
    return max(eligible, key=lambda c: c.area)


def select_chat_by_title(
    candidates: list[EditCandidate],
    title_needle: str,
    min_area: int = MIN_CHAT_AREA,
    min_height: int = MIN_CHAT_HEIGHT,
) -> EditCandidate:
    needle = (title_needle or "").strip()
    if not needle:
        raise RichEditReaderError("chat window title is required")
    matched = [
        c for c in candidates if c.visible and title_matches(c.parent_title, needle)
    ]
    parent_hwnds = {c.parent_hwnd for c in matched}
    if not parent_hwnds:
        raise RichEditReaderError(
            f"no visible {CHAT_EDIT_CLASS} whose window title contains {needle!r}"
        )
    if len(parent_hwnds) > 1:
        titles = sorted({c.parent_title for c in matched})
        raise RichEditReaderError(
            f"ambiguous chat windows matching {needle!r}: {titles}"
        )
    chosen = pick_chat_edit(matched, min_area=min_area, min_height=min_height)
    if chosen is None:
        raise RichEditReaderError(
            f"no visible {CHAT_EDIT_CLASS} chat pane in window matching {needle!r}"
        )
    return chosen


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


def find_chat_richedit(process_name: str, title_needle: str) -> tuple[int, int]:
    candidates = list_richedit_candidates(process_name)
    chosen = select_chat_by_title(candidates, title_needle)
    return chosen.parent_hwnd, chosen.hwnd


def _pack_textrange(cp_min: int, cp_max: int, remote_text_ptr: int, is_32bit: bool) -> bytes:
    if is_32bit:
        return struct.pack("<iiI", cp_min, cp_max, remote_text_ptr & 0xFFFFFFFF)
    return struct.pack("<iiQ", cp_min, cp_max, remote_text_ptr)


def read_richedit_tail(hwnd: int, char_len: int, max_chars: int = MAX_TEXT_CHARS) -> str:
    cp_min = max(0, int(char_len) - int(max_chars))
    cp_max = int(char_len)
    count = cp_max - cp_min
    if count <= 0:
        return ""
    _, pid = win32process.GetWindowThreadProcessId(hwnd)
    process = kernel32.OpenProcess(PROCESS_ACCESS, False, int(pid))
    if not process:
        raise RichEditReaderError(f"OpenProcess failed for pid={pid}")
    try:
        try:
            is_32bit = process_is_32bit(process)
        except Exception as exc:
            raise RichEditReaderError(f"IsWow64Process failed: {exc}") from exc
        text_bytes = (count + 1) * 2
        header = _pack_textrange(cp_min, cp_max, 0, is_32bit)
        remote = kernel32.VirtualAllocEx(
            process,
            None,
            len(header) + text_bytes,
            MEM_COMMIT | MEM_RESERVE,
            PAGE_READWRITE,
        )
        if not remote:
            raise RichEditReaderError("VirtualAllocEx failed")
        try:
            remote_i = ctypes.cast(remote, ctypes.c_void_p).value or 0
            remote_text = remote_i + len(header)
            header = _pack_textrange(cp_min, cp_max, remote_text, is_32bit)
            written = ctypes.c_size_t(0)
            if not kernel32.WriteProcessMemory(
                process,
                remote,
                header,
                len(header),
                ctypes.byref(written),
            ):
                raise RichEditReaderError("WriteProcessMemory failed")
            zero = b"\x00" * text_bytes
            if not kernel32.WriteProcessMemory(
                process,
                ctypes.c_void_p(remote_text),
                zero,
                len(zero),
                ctypes.byref(written),
            ):
                raise RichEditReaderError("WriteProcessMemory text buffer failed")
            user32.SendMessageW(hwnd, EM_GETTEXTRANGE, 0, remote_i)
            buf = (ctypes.c_char * text_bytes)()
            read = ctypes.c_size_t(0)
            if not kernel32.ReadProcessMemory(
                process,
                ctypes.c_void_p(remote_text),
                buf,
                text_bytes,
                ctypes.byref(read),
            ):
                raise RichEditReaderError("ReadProcessMemory failed")
            return bytes(buf).decode("utf-16-le", errors="ignore").split("\x00", 1)[0]
        finally:
            kernel32.VirtualFreeEx(process, remote, 0, MEM_RELEASE)
    finally:
        kernel32.CloseHandle(process)


def read_richedit_text(hwnd: int, max_chars: int = MAX_TEXT_CHARS) -> tuple[str, int, bool]:
    try:
        n = read_richedit_length(hwnd)
    except Exception as exc:
        raise RichEditReaderError(f"WM_GETTEXTLENGTH failed: {exc}") from exc
    if n <= 0:
        return "", 0, False
    if is_gettext_clipped(n, max_chars):
        return read_richedit_tail(hwnd, n, max_chars=max_chars), n, True
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
    return RichEditSnapshot(
        lines=normalize_lines([text]),
        char_len=char_len,
        clipped=clipped,
    )


def read_richedit_lines(hwnd: int) -> list[str]:
    snap = read_richedit_snapshot(hwnd)
    return snap.lines
