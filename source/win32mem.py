from __future__ import annotations

import ctypes
from ctypes import wintypes

import psutil

PROCESS_VM_OPERATION = 0x0008
PROCESS_VM_READ = 0x0010
PROCESS_VM_WRITE = 0x0020
PROCESS_QUERY_INFORMATION = 0x0400
PROCESS_ACCESS = (
    PROCESS_VM_OPERATION
    | PROCESS_VM_READ
    | PROCESS_VM_WRITE
    | PROCESS_QUERY_INFORMATION
)

MEM_COMMIT = 0x1000
MEM_RESERVE = 0x2000
MEM_RELEASE = 0x8000
PAGE_READWRITE = 0x04

kernel32 = ctypes.WinDLL("kernel32", use_last_error=True)
kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
kernel32.OpenProcess.restype = wintypes.HANDLE
kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
kernel32.CloseHandle.restype = wintypes.BOOL
kernel32.VirtualAllocEx.argtypes = [
    wintypes.HANDLE,
    wintypes.LPVOID,
    ctypes.c_size_t,
    wintypes.DWORD,
    wintypes.DWORD,
]
kernel32.VirtualAllocEx.restype = wintypes.LPVOID
kernel32.VirtualFreeEx.argtypes = [
    wintypes.HANDLE,
    wintypes.LPVOID,
    ctypes.c_size_t,
    wintypes.DWORD,
]
kernel32.VirtualFreeEx.restype = wintypes.BOOL
kernel32.WriteProcessMemory.argtypes = [
    wintypes.HANDLE,
    wintypes.LPVOID,
    wintypes.LPCVOID,
    ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_size_t),
]
kernel32.WriteProcessMemory.restype = wintypes.BOOL
kernel32.ReadProcessMemory.argtypes = [
    wintypes.HANDLE,
    wintypes.LPCVOID,
    wintypes.LPVOID,
    ctypes.c_size_t,
    ctypes.POINTER(ctypes.c_size_t),
]
kernel32.ReadProcessMemory.restype = wintypes.BOOL
kernel32.IsWow64Process.argtypes = [wintypes.HANDLE, ctypes.POINTER(wintypes.BOOL)]
kernel32.IsWow64Process.restype = wintypes.BOOL


def normalize_lines(chunks: list[str]) -> list[str]:
    lines: list[str] = []
    seen: set[str] = set()
    for chunk in chunks:
        for line in str(chunk).splitlines():
            cleaned = line.strip()
            if not cleaned or cleaned in seen:
                continue
            seen.add(cleaned)
            lines.append(cleaned)
    return lines


def process_pids(process_name: str) -> set[int]:
    expected = process_name.lower()
    pids: set[int] = set()
    for proc in psutil.process_iter(["name", "pid"]):
        try:
            if (proc.info.get("name") or "").lower() == expected:
                pids.add(int(proc.info["pid"]))
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
    return pids


def process_is_32bit(process: wintypes.HANDLE) -> bool:
    if ctypes.sizeof(ctypes.c_void_p) == 4:
        return True
    wow64 = wintypes.BOOL()
    if not kernel32.IsWow64Process(process, ctypes.byref(wow64)):
        raise RuntimeError("IsWow64Process failed")
    return bool(wow64.value)
