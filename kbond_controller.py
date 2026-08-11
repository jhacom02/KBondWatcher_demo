"""K-Bond (axis.exe) HWND discovery, relative click, and clipboard paste."""

from __future__ import annotations

import logging
import time
from dataclasses import dataclass
from typing import Optional

import psutil
import win32api
import win32clipboard
import win32con
import win32gui
import win32process

from models import KBondWindowInfo

logger = logging.getLogger("forestbond_watcher")


class KBondError(RuntimeError):
    """Raised when K-Bond process/window automation fails."""


@dataclass(frozen=True)
class ClientClickPoint:
    client_x: int
    client_y: int
    screen_x: int
    screen_y: int
    client_width: int
    client_height: int


def compute_client_click(
    client_width: int,
    client_height: int,
    win_x: float,
    win_y: float,
) -> tuple[int, int]:
    """Pure helper: client-relative pixel from ratios (for unit tests)."""
    if client_width < 0 or client_height < 0:
        raise ValueError("client dimensions must be non-negative")
    if not (0.0 <= win_x <= 1.0 and 0.0 <= win_y <= 1.0):
        raise ValueError("win_x/win_y must be in [0, 1]")
    return int(client_width * win_x), int(client_height * win_y)


def resolve_kbond_pid(process_name: str, configured_pid: int) -> int:
    """Return a single validated PID for the K-Bond process."""
    expected = process_name.lower()
    if configured_pid > 0:
        try:
            proc = psutil.Process(configured_pid)
        except psutil.NoSuchProcess as exc:
            raise KBondError(
                f"KBOND_PID={configured_pid} does not exist. "
                f"Start {process_name} or update KBOND_PID."
            ) from exc
        except psutil.AccessDenied as exc:
            raise KBondError(
                f"Access denied reading process KBOND_PID={configured_pid}."
            ) from exc

        actual = (proc.name() or "").lower()
        if actual != expected:
            raise KBondError(
                f"KBOND_PID={configured_pid} is '{proc.name()}', "
                f"expected '{process_name}'."
            )
        if not proc.is_running():
            raise KBondError(f"KBOND_PID={configured_pid} is not running.")
        return configured_pid

    matches: list[int] = []
    for proc in psutil.process_iter(["pid", "name"]):
        try:
            name = (proc.info.get("name") or "").lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        if name == expected:
            matches.append(int(proc.info["pid"]))

    if not matches:
        raise KBondError(
            f"No process named '{process_name}' found. "
            "Start K-Bond or set KBOND_PID."
        )
    if len(matches) > 1:
        raise KBondError(
            f"Multiple {process_name} processes found: {matches}. "
            "Set KBOND_PID to the correct PID."
        )
    return matches[0]


def _window_area(hwnd: int) -> int:
    left, top, right, bottom = win32gui.GetWindowRect(hwnd)
    return max(0, right - left) * max(0, bottom - top)


# Candidate: (hwnd, title, area, pid)
WindowCandidate = tuple[int, str, int, int]


_SHELL_PROCESS_BLOCKLIST = frozenset(
    {
        "explorer.exe",
        "dwm.exe",
        "shellexperiencehost.exe",
        "searchhost.exe",
        "searchapp.exe",
        "startmenuexperiencehost.exe",
        "applicationframehost.exe",
        "systemsettings.exe",
        "textinputhost.exe",
    }
)


def _process_name_tokens(process_name: str) -> set[str]:
    """Tokens used to find helper/login executables related to K-Bond."""
    tokens: set[str] = {"axis", "kbond", "k-bond", "k_bond"}
    name = (process_name or "").lower().strip()
    if name:
        tokens.add(name)
        if name.endswith(".exe"):
            tokens.add(name[:-4])
        else:
            tokens.add(f"{name}.exe")
    # Keep only reasonably specific tokens.
    return {t for t in tokens if len(t) >= 4}


def collect_related_pids(root_pid: int, process_name: str = "") -> set[int]:
    """
    Expand PID scope for login/helper windows.

    Includes: root, recursive children, same exe name, and fuzzy name matches
    (e.g. AxisLogin.exe when KBOND_PROCESS_NAME=axis.exe).
    Does NOT include parent/siblings (those are often explorer.exe).
    """
    related: set[int] = {int(root_pid)}
    expected = (process_name or "").lower()
    tokens = _process_name_tokens(expected)

    try:
        root = psutil.Process(root_pid)
        for child in root.children(recursive=True):
            related.add(int(child.pid))
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        pass

    for proc in psutil.process_iter(["pid", "name"]):
        try:
            name = (proc.info.get("name") or "").lower()
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            continue
        if not name or name in _SHELL_PROCESS_BLOCKLIST:
            continue
        if expected and name == expected:
            related.add(int(proc.info["pid"]))
            continue
        base = name[:-4] if name.endswith(".exe") else name
        if any(token in base for token in tokens):
            related.add(int(proc.info["pid"]))

    return related


def _process_name_for_pid(pid: int) -> str:
    try:
        return (psutil.Process(pid).name() or "").lower()
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        return ""


def _is_blocked_shell_candidate(pid: int, title: str) -> bool:
    name = _process_name_for_pid(pid)
    if name in _SHELL_PROCESS_BLOCKLIST:
        return True
    # Desktop / Progman must never be treated as K-Bond.
    if title.strip().lower() in {"program manager", "progman"}:
        return True
    return False


def _is_enumerable_window(hwnd: int, *, allow_hidden: bool = False) -> bool:
    """Accept top-level windows; optionally include hidden ones with non-zero area."""
    if not win32gui.IsWindow(hwnd):
        return False
    if win32gui.GetParent(hwnd) != 0:
        return False
    if win32gui.IsWindowVisible(hwnd) or win32gui.IsIconic(hwnd):
        return True
    if allow_hidden:
        try:
            return _window_area(hwnd) > 0
        except Exception:  # noqa: BLE001
            return False
    return False


def enumerate_visible_candidates(
    allowed_pids: Optional[set[int]] = None,
    *,
    allow_hidden: bool = False,
) -> list[WindowCandidate]:
    """Enumerate top-level windows, optionally filtered by PID set."""
    candidates: list[WindowCandidate] = []

    def _enum(hwnd: int, _: object) -> None:
        if not _is_enumerable_window(hwnd, allow_hidden=allow_hidden):
            return
        _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
        window_pid = int(window_pid)
        if allowed_pids is not None and window_pid not in allowed_pids:
            return
        title = win32gui.GetWindowText(hwnd) or ""
        if _is_blocked_shell_candidate(window_pid, title):
            return
        area = _window_area(hwnd)
        if area <= 0:
            return
        candidates.append((int(hwnd), title, area, window_pid))

    win32gui.EnumWindows(_enum, None)
    return candidates


def select_best_hwnd(
    candidates: list[WindowCandidate],
    title_contains: str = "",
    preferred_pid: Optional[int] = None,
    require_title: bool = False,
) -> Optional[int]:
    """Pick HWND: title match > preferred PID > largest area."""
    if not candidates:
        return None

    pool = list(candidates)
    needle = (title_contains or "").strip().lower()
    if needle:
        titled = [c for c in pool if needle in c[1].lower()]
        if titled:
            pool = titled
        elif require_title:
            return None

    if preferred_pid is not None:
        preferred = [c for c in pool if c[3] == int(preferred_pid)]
        if preferred:
            pool = preferred

    pool.sort(key=lambda item: item[2], reverse=True)
    return pool[0][0]


def find_windows_by_title_keywords(
    keywords: list[str],
) -> list[WindowCandidate]:
    """Find visible windows whose title contains any keyword (case-insensitive)."""
    keys = [k.strip().lower() for k in keywords if k and k.strip()]
    if not keys:
        return []
    all_windows = enumerate_visible_candidates(None)
    matched: list[WindowCandidate] = []
    for cand in all_windows:
        title_l = cand[1].lower()
        if any(k in title_l for k in keys):
            matched.append(cand)
    return matched


def build_title_keywords(title_contains: str = "") -> list[str]:
    keywords: list[str] = []
    if (title_contains or "").strip():
        keywords.append(title_contains.strip())
    for alias in ("KBond", "K-Bond", "KBOND", "로그인"):
        if alias.lower() not in {k.lower() for k in keywords}:
            keywords.append(alias)
    return keywords


def scan_kbond_candidates(
    root_pid: int,
    process_name: str = "",
    title_contains: str = "",
) -> str:
    """Human-readable scan used by --diagnose-kbond (even when selection fails)."""
    related = sorted(collect_related_pids(root_pid, process_name=process_name))
    lines: list[str] = [
        f"root_pid={root_pid}",
        f"process_name={process_name!r}",
        f"title_contains={title_contains!r}",
        f"related_pids={related}",
        "---- processes ----",
    ]
    for pid in related:
        lines.append(f"  pid={pid} name={_process_name_for_pid(pid)!r}")

    visible = enumerate_visible_candidates(set(related), allow_hidden=False)
    hidden = enumerate_visible_candidates(set(related), allow_hidden=True)
    lines.append(f"---- windows on related PIDs (visible={len(visible)}) ----")
    for hwnd, title, area, pid in sorted(visible, key=lambda c: c[2], reverse=True):
        lines.append(
            f"  hwnd=0x{hwnd:08X} pid={pid} area={area} title={title!r} "
            f"proc={_process_name_for_pid(pid)!r}"
        )
    extra_hidden = [c for c in hidden if c not in visible]
    if extra_hidden:
        lines.append(f"---- hidden/non-visible with area (n={len(extra_hidden)}) ----")
        for hwnd, title, area, pid in sorted(extra_hidden, key=lambda c: c[2], reverse=True)[:30]:
            lines.append(
                f"  hwnd=0x{hwnd:08X} pid={pid} area={area} title={title!r} "
                f"proc={_process_name_for_pid(pid)!r}"
            )

    keywords = build_title_keywords(title_contains)
    titled = find_windows_by_title_keywords(keywords)
    lines.append(f"---- title keyword matches {keywords!r} (n={len(titled)}) ----")
    for hwnd, title, area, pid in sorted(titled, key=lambda c: c[2], reverse=True)[:40]:
        lines.append(
            f"  hwnd=0x{hwnd:08X} pid={pid} area={area} title={title!r} "
            f"proc={_process_name_for_pid(pid)!r}"
        )
    return "\n".join(lines)


def find_top_level_hwnd(
    pid: int,
    title_contains: str = "",
    process_name: str = "",
) -> int:
    """
    Find a usable K-Bond HWND (main or login).

    Search order:
    1) Exact configured PID (visible, then hidden-with-area)
    2) Related PIDs: children + same/fuzzy exe names (axis/kbond*)
    3) Title keywords across all processes
    """
    root_pid = int(pid)
    expected_name = (process_name or "").lower()
    related_pids = collect_related_pids(root_pid, process_name=process_name)

    def _pick(
        candidates: list[WindowCandidate],
        *,
        require_title: bool = False,
        label: str,
    ) -> Optional[int]:
        hwnd = select_best_hwnd(
            candidates,
            title_contains=title_contains,
            preferred_pid=root_pid,
            require_title=require_title,
        )
        if hwnd is None:
            return None
        match = next(c for c in candidates if c[0] == hwnd)
        logger.info(
            "%s | root_pid=%s window_pid=%s title=%r proc=%s",
            label,
            root_pid,
            match[3],
            match[1],
            _process_name_for_pid(match[3]),
        )
        return hwnd

    # 1) Exact PID
    exact = enumerate_visible_candidates({root_pid})
    hwnd = _pick(exact, label="KBOND_HWND_EXACT")
    if hwnd is not None:
        return hwnd
    exact_hidden = enumerate_visible_candidates({root_pid}, allow_hidden=True)
    hwnd = _pick(exact_hidden, label="KBOND_HWND_EXACT_HIDDEN")
    if hwnd is not None:
        return hwnd

    # 2) Related fuzzy processes (login helpers often use another exe name)
    related = enumerate_visible_candidates(related_pids)
    if expected_name:
        same_exe = [
            c for c in related if _process_name_for_pid(c[3]) == expected_name
        ]
        hwnd = _pick(same_exe, label="KBOND_HWND_SAME_EXE")
        if hwnd is not None:
            return hwnd
    hwnd = _pick(related, label="KBOND_HWND_RELATED")
    if hwnd is not None:
        return hwnd

    related_hidden = enumerate_visible_candidates(related_pids, allow_hidden=True)
    hwnd = _pick(related_hidden, label="KBOND_HWND_RELATED_HIDDEN")
    if hwnd is not None:
        return hwnd

    # 3) Title keyword fallback
    keywords = build_title_keywords(title_contains)
    titled = find_windows_by_title_keywords(keywords)
    if (title_contains or "").strip():
        hwnd = _pick(titled, require_title=True, label="KBOND_HWND_TITLE")
        if hwnd is not None:
            return hwnd
    hwnd = _pick(titled, label="KBOND_HWND_TITLE_ALIAS")
    if hwnd is not None:
        return hwnd

    raise KBondError(
        "No visible window found for K-Bond/login.\n"
        + scan_kbond_candidates(
            root_pid,
            process_name=process_name,
            title_contains=title_contains,
        )
    )


def get_click_point(hwnd: int, win_x: float, win_y: float) -> ClientClickPoint:
    left, top, right, bottom = win32gui.GetClientRect(hwnd)
    width = right - left
    height = bottom - top
    if width <= 0 or height <= 0:
        raise KBondError(
            f"Invalid ClientRect for HWND=0x{hwnd:08X}: "
            f"width={width}, height={height}"
        )

    client_x, client_y = compute_client_click(width, height, win_x, win_y)
    screen_x, screen_y = win32gui.ClientToScreen(hwnd, (client_x, client_y))
    return ClientClickPoint(
        client_x=client_x,
        client_y=client_y,
        screen_x=int(screen_x),
        screen_y=int(screen_y),
        client_width=width,
        client_height=height,
    )


def restore_and_foreground(hwnd: int) -> None:
    if win32gui.IsIconic(hwnd):
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        time.sleep(0.15)

    try:
        win32gui.SetForegroundWindow(hwnd)
    except Exception as exc:  # noqa: BLE001 - foreground often blocked
        logger.warning(
            "SetForegroundWindow failed for HWND=0x%08X: %s "
            "(continuing with click attempt)",
            hwnd,
            exc,
        )


def click_screen(x: int, y: int, settle_ms: int = 150) -> None:
    win32api.SetCursorPos((int(x), int(y)))
    time.sleep(0.05)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    time.sleep(max(settle_ms, 0) / 1000.0)


def _get_clipboard_text() -> Optional[str]:
    try:
        win32clipboard.OpenClipboard()
        try:
            if win32clipboard.IsClipboardFormatAvailable(win32con.CF_UNICODETEXT):
                data = win32clipboard.GetClipboardData(win32con.CF_UNICODETEXT)
                return str(data) if data is not None else ""
            return None
        finally:
            win32clipboard.CloseClipboard()
    except Exception:  # noqa: BLE001
        return None


def _set_clipboard_text(text: str) -> None:
    win32clipboard.OpenClipboard()
    try:
        win32clipboard.EmptyClipboard()
        win32clipboard.SetClipboardData(win32con.CF_UNICODETEXT, text)
    finally:
        win32clipboard.CloseClipboard()


def paste_text_via_clipboard(text: str) -> None:
    """Paste Unicode text with Ctrl+V to avoid Hangul IME typing issues."""
    from pywinauto.keyboard import send_keys

    previous = _get_clipboard_text()
    try:
        _set_clipboard_text(text)
        time.sleep(0.05)
        send_keys("^v")
        time.sleep(0.1)
    finally:
        try:
            if previous is None:
                win32clipboard.OpenClipboard()
                try:
                    win32clipboard.EmptyClipboard()
                finally:
                    win32clipboard.CloseClipboard()
            else:
                _set_clipboard_text(previous)
        except Exception as exc:  # noqa: BLE001
            logger.warning("Failed to restore clipboard: %s", exc)


def inspect_kbond(
    process_name: str,
    configured_pid: int,
    win_x: float,
    win_y: float,
    title_contains: str = "",
) -> KBondWindowInfo:
    pid = resolve_kbond_pid(process_name, configured_pid)
    hwnd = find_top_level_hwnd(
        pid,
        title_contains=title_contains,
        process_name=process_name,
    )
    title = win32gui.GetWindowText(hwnd) or ""
    _, window_pid = win32process.GetWindowThreadProcessId(hwnd)
    window_pid = int(window_pid)
    window_rect = tuple(win32gui.GetWindowRect(hwnd))
    client_rect = tuple(win32gui.GetClientRect(hwnd))
    click = get_click_point(hwnd, win_x, win_y)

    try:
        actual_name = psutil.Process(window_pid).name()
    except Exception:  # noqa: BLE001
        actual_name = process_name

    if window_pid != pid:
        logger.info(
            "KBOND_PID_WINDOW | configured/root=%s window=%s (%s)",
            pid,
            window_pid,
            actual_name,
        )

    return KBondWindowInfo(
        pid=window_pid,
        process_name=actual_name,
        hwnd=int(hwnd),
        title=title,
        window_rect=(
            int(window_rect[0]),
            int(window_rect[1]),
            int(window_rect[2]),
            int(window_rect[3]),
        ),
        client_rect=(
            int(client_rect[0]),
            int(client_rect[1]),
            int(client_rect[2]),
            int(client_rect[3]),
        ),
        click_client=(click.client_x, click.client_y),
        click_screen=(click.screen_x, click.screen_y),
    )


def prefill_kbond(
    process_name: str,
    configured_pid: int,
    win_x: float,
    win_y: float,
    send_text: str,
    title_contains: str = "",
) -> KBondWindowInfo:
    """
    Click relative client coordinate and paste send_text.

    Never sends Enter for axis.exe / real K-Bond. Caller must not submit.
    """
    info = inspect_kbond(
        process_name=process_name,
        configured_pid=configured_pid,
        win_x=win_x,
        win_y=win_y,
        title_contains=title_contains,
    )
    logger.info("KBOND_PID | %s", info.pid)
    logger.info("KBOND_HWND | 0x%08X", info.hwnd)
    logger.info(
        "KBOND_RECT | window=%s client=%s",
        info.window_rect,
        info.client_rect,
    )

    restore_and_foreground(info.hwnd)
    click_screen(info.click_screen[0], info.click_screen[1])
    logger.info("CLICK | x=%s, y=%s", info.click_screen[0], info.click_screen[1])

    paste_text_via_clipboard(send_text)
    logger.info('PREFILLED | "%s"', send_text)
    # Intentionally no Enter / WM_KEYDOWN / UIA Invoke submit.
    return info


def format_diagnose_report(info: KBondWindowInfo, win_x: float, win_y: float) -> str:
    lines = [
        f"PID: {info.pid}",
        f"process name: {info.process_name}",
        f"HWND: 0x{info.hwnd:08X} ({info.hwnd})",
        f"window title: {info.title!r}",
        f"WindowRect: {info.window_rect}",
        f"ClientRect: {info.client_rect}",
        f"WIN_X/WIN_Y: {win_x} / {win_y}",
        f"click client: {info.click_client}",
        f"click screen: {info.click_screen}",
    ]
    return "\n".join(lines)
