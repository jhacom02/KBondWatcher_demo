from __future__ import annotations

import ctypes
import struct
from ctypes import wintypes
from dataclasses import dataclass
from typing import Optional

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

TREE_CLASS = "TElTree"
CHAT_MIN_CENTER_X_RATIO = 0.55

TVM_GETCOUNT = 0x1105
TVM_GETNEXTITEM = 0x110A
TVM_GETITEMA = 0x110C
TVM_GETITEMW = 0x113E

TVGN_ROOT = 0x0000
TVGN_NEXT = 0x0001
TVGN_CHILD = 0x0004

TVIF_TEXT = 0x0001

TEXT_BUFFER_CHARS = 512
MAX_ITEMS = 5000

user32 = ctypes.WinDLL("user32", use_last_error=True)
user32.SendMessageW.argtypes = [
    wintypes.HWND,
    wintypes.UINT,
    wintypes.WPARAM,
    wintypes.LPARAM,
]
user32.SendMessageW.restype = wintypes.LPARAM


class ElTreeReaderError(RuntimeError):
    pass


@dataclass(frozen=True)
class TreeCandidate:
    hwnd: int
    left: int
    top: int
    right: int
    bottom: int

    @property
    def area(self) -> int:
        return max(0, self.right - self.left) * max(0, self.bottom - self.top)

    def center_x_ratio(self, parent_left: int, parent_right: int) -> float:
        width = parent_right - parent_left
        if width <= 0:
            return 0.0
        center = (self.left + self.right) / 2.0
        return (center - parent_left) / width


def pick_chat_tree(
    candidates: list[TreeCandidate],
    parent_left: int,
    parent_right: int,
    min_center_x_ratio: float = CHAT_MIN_CENTER_X_RATIO,
) -> Optional[TreeCandidate]:
    eligible = [
        c
        for c in candidates
        if c.center_x_ratio(parent_left, parent_right) >= min_center_x_ratio
    ]
    if not eligible:
        return None
    return max(eligible, key=lambda c: c.area)


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


def find_messenger_hwnd(process_name: str, title_needle: str) -> int:
    needle = title_needle.lower()
    pids = process_pids(process_name)
    if not pids:
        raise ElTreeReaderError(f"process not running: {process_name!r}")

    matches: list[int] = []

    def _callback(hwnd: int, _: object) -> bool:
        try:
            title = win32gui.GetWindowText(hwnd) or ""
            if needle not in title.lower():
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if int(pid) not in pids:
                return True
            matches.append(hwnd)
        except Exception:
            pass
        return True

    win32gui.EnumWindows(_callback, None)
    if not matches:
        raise ElTreeReaderError(
            f"window containing title {title_needle!r} not found "
            f"(process={process_name!r})"
        )
    return max(matches, key=_rank_window)


def list_eltree_candidates(parent_hwnd: int) -> list[TreeCandidate]:
    found: list[TreeCandidate] = []

    def _callback(hwnd: int, _: object) -> bool:
        try:
            cls = win32gui.GetClassName(hwnd) or ""
            if cls != TREE_CLASS:
                return True
            if not win32gui.IsWindowVisible(hwnd):
                return True
            left, top, right, bottom = win32gui.GetWindowRect(hwnd)
            found.append(
                TreeCandidate(
                    hwnd=int(hwnd),
                    left=left,
                    top=top,
                    right=right,
                    bottom=bottom,
                )
            )
        except Exception:
            pass
        return True

    win32gui.EnumChildWindows(parent_hwnd, _callback, None)
    return found


def find_chat_eltree(parent_hwnd: int) -> int:
    parent_left, _pt, parent_right, _pb = win32gui.GetWindowRect(parent_hwnd)
    candidates = list_eltree_candidates(parent_hwnd)
    chosen = pick_chat_tree(candidates, parent_left, parent_right)
    if chosen is None:
        raise ElTreeReaderError(
            f"no visible {TREE_CLASS} in chat region "
            f"(candidates={len(candidates)})"
        )
    return chosen.hwnd


def _pack_tvitem_text(
    *,
    hitem: int,
    remote_text_ptr: int,
    cch: int,
    is_32bit: bool,
) -> bytes:
    mask = TVIF_TEXT
    if is_32bit:
        return struct.pack(
            "<IIIIIiiiiI",
            mask,
            hitem & 0xFFFFFFFF,
            0,
            0,
            remote_text_ptr & 0xFFFFFFFF,
            cch,
            0,
            0,
            0,
            0,
        )
    return struct.pack(
        "<I4xQIIQiiiiQ",
        mask,
        hitem,
        0,
        0,
        remote_text_ptr,
        cch,
        0,
        0,
        0,
        0,
    )


def _send(hwnd: int, msg: int, wparam: int, lparam: int) -> int:
    return int(user32.SendMessageW(hwnd, msg, wparam, lparam))


def _read_item_text(
    process: wintypes.HANDLE,
    tree_hwnd: int,
    hitem: int,
    is_32bit: bool,
    unicode: bool,
) -> str:
    char_bytes = 2 if unicode else 1
    text_bytes = TEXT_BUFFER_CHARS * char_bytes
    item_blob = _pack_tvitem_text(
        hitem=hitem,
        remote_text_ptr=0,
        cch=TEXT_BUFFER_CHARS,
        is_32bit=is_32bit,
    )
    remote = kernel32.VirtualAllocEx(
        process,
        None,
        len(item_blob) + text_bytes,
        MEM_COMMIT | MEM_RESERVE,
        PAGE_READWRITE,
    )
    if not remote:
        raise ElTreeReaderError("VirtualAllocEx failed")
    try:
        remote_i = ctypes.cast(remote, ctypes.c_void_p).value or 0
        remote_text = remote_i + len(item_blob)
        item_blob = _pack_tvitem_text(
            hitem=hitem,
            remote_text_ptr=remote_text,
            cch=TEXT_BUFFER_CHARS,
            is_32bit=is_32bit,
        )
        written = ctypes.c_size_t(0)
        if not kernel32.WriteProcessMemory(
            process,
            remote,
            item_blob,
            len(item_blob),
            ctypes.byref(written),
        ):
            raise ElTreeReaderError("WriteProcessMemory failed")
        zero = b"\x00" * text_bytes
        if not kernel32.WriteProcessMemory(
            process,
            ctypes.c_void_p(remote_text),
            zero,
            len(zero),
            ctypes.byref(written),
        ):
            raise ElTreeReaderError("WriteProcessMemory text buffer failed")
        msg = TVM_GETITEMW if unicode else TVM_GETITEMA
        _send(tree_hwnd, msg, 0, remote_i)
        buf = (ctypes.c_char * text_bytes)()
        read = ctypes.c_size_t(0)
        if not kernel32.ReadProcessMemory(
            process,
            ctypes.c_void_p(remote_text),
            buf,
            text_bytes,
            ctypes.byref(read),
        ):
            raise ElTreeReaderError("ReadProcessMemory failed")
        raw = bytes(buf)
        if unicode:
            return raw.decode("utf-16-le", errors="ignore").split("\x00", 1)[0]
        return raw.split(b"\x00", 1)[0].decode("mbcs", errors="ignore")
    finally:
        kernel32.VirtualFreeEx(process, remote, 0, MEM_RELEASE)


def _iter_item_handles(tree_hwnd: int) -> list[int]:
    items: list[int] = []
    stack: list[int] = []
    item = _send(tree_hwnd, TVM_GETNEXTITEM, TVGN_ROOT, 0)
    while item and len(items) < MAX_ITEMS:
        items.append(item)
        child = _send(tree_hwnd, TVM_GETNEXTITEM, TVGN_CHILD, item)
        nxt = _send(tree_hwnd, TVM_GETNEXTITEM, TVGN_NEXT, item)
        if child:
            if nxt:
                stack.append(nxt)
            item = child
        elif nxt:
            item = nxt
        elif stack:
            item = stack.pop()
        else:
            break
    return items


def read_eltree_lines(tree_hwnd: int) -> list[str]:
    if not tree_hwnd or not win32gui.IsWindow(tree_hwnd):
        raise ElTreeReaderError("invalid TElTree hwnd")
    _, pid = win32process.GetWindowThreadProcessId(tree_hwnd)
    process = kernel32.OpenProcess(PROCESS_ACCESS, False, int(pid))
    if not process:
        raise ElTreeReaderError(f"OpenProcess failed for pid={pid}")
    try:
        is_32bit = process_is_32bit(process)
        count = _send(tree_hwnd, TVM_GETCOUNT, 0, 0)
        handles = _iter_item_handles(tree_hwnd)
        if count <= 0 and not handles:
            raise ElTreeReaderError("TElTree item count is 0")
        unicode = bool(win32gui.IsWindowUnicode(tree_hwnd))
        chunks: list[str] = []
        for hitem in handles:
            text = _read_item_text(
                process,
                tree_hwnd,
                hitem,
                is_32bit=is_32bit,
                unicode=unicode,
            )
            if text:
                chunks.append(text)
        lines = normalize_lines(chunks)
        if not lines:
            raise ElTreeReaderError("TElTree returned no text lines")
        return lines
    finally:
        kernel32.CloseHandle(process)


def item_count(tree_hwnd: int) -> int:
    return _send(tree_hwnd, TVM_GETCOUNT, 0, 0)
