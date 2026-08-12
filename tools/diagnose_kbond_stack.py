#!/usr/bin/env python3
"""Offline/login-optional stack hints for K-Bond install + running windows.

Does not import KBondWatcher app modules. Safe to run without login.
Chat-body class / UIA text exposure still need the messenger UI after login.
"""

from __future__ import annotations

import argparse
import struct
import sys
from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional


DEFAULT_ROOTS = (
    Path(r"C:\KBond"),
    Path(r"C:\KBond\exe"),
    Path(r"C:\KBond\KBondMessenger"),
)

INTERESTING_NAMES = {
    "axis.exe",
    "axisver.exe",
    "axiscode.dll",
    "axiscodx.dll",
    "axisdialog.dll",
    "axisform.dll",
    "axislib.dll",
    "axislogin.dll",
    "axissm.dll",
    "axisvbs.dll",
    "axsock.ocx",
    "axwizard.ocx",
    "axxecure.ocx",
    "kbondmessenger.exe",
    "launcher.exe",
    "launcherctrl.exe",
    "messengercore.dll",
    "mss.dll",
    "mfc100.dll",
    "mfc100u.dll",
    "msvcr100.dll",
}

PROCESS_NEEDLES_EXACT = {
    "axis.exe",
    "axisver.exe",
    "kbondmessenger.exe",
    "launcher.exe",
    "launcherctrl.exe",
    "ezmailchecker.exe",
    "ezlanguagechanger.exe",
    "ezwebplugin.exe",
}

DELPHI_STRINGS = (
    b"Borland",
    b"Embarcadero",
    b"SysInit",
    b"System.SysUtils",
    b"Vcl.Forms",
    b"Vcl.Controls",
    b"TApplication",
    b"TElTree",
    b"ElTree",
    b"ElPack",
    b"PACKAGEINFO",
    b"@System@",
    b"rtl120.bpl",
    b"rtl140.bpl",
    b"rtl160.bpl",
    b"rtl170.bpl",
    b"vcl120.bpl",
    b"vcl140.bpl",
    b"vcl160.bpl",
    b"This program must be run under Win32",
)

MFC_STRINGS = (
    b"MFC100",
    b"MFC140",
    b"AfxWinMain",
    b"AfxGetApp",
    b"CWinApp",
    b"CDialog",
    b"mfc100u.dll",
    b"mfc140u.dll",
)

QT_STRINGS = (b"Qt5Core", b"Qt6Core", b"QWidget", b"QtWebEngine")
ELECTRON_STRINGS = (b"Electron", b"chrome_elf.dll", b"resources.pak")
DOTNET_STRINGS = (b"mscoree.dll", b"BSJB", b".NETFramework")

CLASS_HINTS = {
    "teltree": "ElTree (Delphi ElPack) - strong if on chat body",
    "eltree": "ElTree-related",
    "tform": "Delphi VCL form hint",
    "tpanel": "Delphi VCL panel hint",
    "tedit": "Delphi VCL edit hint",
    "tbutton": "Delphi VCL button hint",
    "afx:": "MFC window class hint",
    "afxwnd": "MFC window class hint",
    "#32770": "standard dialog",
}


@dataclass
class FileHint:
    path: Path
    size: int
    delphi_hits: list[str] = field(default_factory=list)
    mfc_hits: list[str] = field(default_factory=list)
    other_hits: list[str] = field(default_factory=list)
    imports: list[str] = field(default_factory=list)
    error: str = ""


def _match_needles(data: bytes, needles: Iterable[bytes]) -> list[str]:
    hits: list[str] = []
    for needle in needles:
        if needle in data:
            try:
                hits.append(needle.decode("ascii"))
            except UnicodeDecodeError:
                hits.append(repr(needle))
    return hits


def _pe_import_dll_names(data: bytes) -> list[str]:
    if len(data) < 0x40 or data[:2] != b"MZ":
        return []
    try:
        e_lfanew = struct.unpack_from("<I", data, 0x3C)[0]
        if data[e_lfanew : e_lfanew + 4] != b"PE\0\0":
            return []
        coff = e_lfanew + 4
        num_sections = struct.unpack_from("<H", data, coff + 2)[0]
        size_opt = struct.unpack_from("<H", data, coff + 16)[0]
        opt = coff + 20
        magic = struct.unpack_from("<H", data, opt)[0]
        if magic == 0x10B:
            import_rva = struct.unpack_from("<I", data, opt + 104)[0]
            pe32_plus = False
        elif magic == 0x20B:
            import_rva = struct.unpack_from("<I", data, opt + 120)[0]
            pe32_plus = True
        else:
            return []
        section_off = opt + size_opt
        sections: list[tuple[int, int, int]] = []
        for i in range(num_sections):
            off = section_off + i * 40
            vsize, va, raw_size, raw_ptr = struct.unpack_from("<IIII", data, off + 8)
            sections.append((va, raw_ptr, max(vsize, raw_size)))

        def rva_to_off(rva: int) -> Optional[int]:
            for va, raw_ptr, size in sections:
                if va <= rva < va + size:
                    return raw_ptr + (rva - va)
            return None

        if not import_rva:
            return []
        names: list[str] = []
        desc_off = rva_to_off(import_rva)
        if desc_off is None:
            return []
        idx = 0
        while True:
            base = desc_off + idx * 20
            if base + 20 > len(data):
                break
            (
                _ilt,
                _td,
                _ft,
                name_rva,
                _iat,
            ) = struct.unpack_from("<IIIII", data, base)
            if name_rva == 0:
                break
            name_off = rva_to_off(name_rva)
            if name_off is None:
                break
            end = data.find(b"\0", name_off)
            raw = data[name_off:end if end != -1 else name_off + 64]
            dll = raw.decode("ascii", errors="ignore").strip()
            if dll:
                names.append(dll)
            idx += 1
            if idx > 512:
                break
        _ = pe32_plus
        return names
    except Exception:
        return []


def analyze_pe(path: Path, max_bytes: int = 12_000_000) -> FileHint:
    hint = FileHint(path=path, size=path.stat().st_size)
    try:
        with path.open("rb") as fh:
            data = fh.read(max_bytes)
    except OSError as exc:
        hint.error = str(exc)
        return hint

    hint.imports = _pe_import_dll_names(data)
    hint.delphi_hits = _match_needles(data, DELPHI_STRINGS)
    hint.mfc_hits = _match_needles(data, MFC_STRINGS)
    other: list[str] = []
    other.extend(f"Qt:{s}" for s in _match_needles(data, QT_STRINGS))
    other.extend(f"Electron:{s}" for s in _match_needles(data, ELECTRON_STRINGS))
    other.extend(f"dotnet:{s}" for s in _match_needles(data, DOTNET_STRINGS))
    for dll in hint.imports:
        low = dll.lower()
        if low.startswith("vcl") or low.startswith("rtl") or "borlndmm" in low:
            hint.delphi_hits.append(f"import:{dll}")
        if low.startswith("mfc") or low.startswith("mfcm"):
            hint.mfc_hits.append(f"import:{dll}")
        if "qt5" in low or "qt6" in low:
            other.append(f"import:{dll}")
    # Unique preserve order
    def uniq(xs: list[str]) -> list[str]:
        seen: set[str] = set()
        out: list[str] = []
        for x in xs:
            if x not in seen:
                seen.add(x)
                out.append(x)
        return out

    hint.delphi_hits = uniq(hint.delphi_hits)
    hint.mfc_hits = uniq(hint.mfc_hits)
    hint.other_hits = uniq(other)
    return hint


def discover_binaries(roots: list[Path]) -> list[Path]:
    found: list[Path] = []
    seen: set[Path] = set()
    for root in roots:
        if not root.exists():
            continue
        if root.is_file():
            candidates = [root]
        else:
            candidates = []
            for pattern in ("*.exe", "*.dll", "*.ocx", "*.bpl"):
                candidates.extend(root.glob(pattern))
            # Limited depth: plugin/Skins often hold large UI dlls worth scanning.
            for sub in ("plugin", "Skins", "Language"):
                subdir = root / sub
                if not subdir.is_dir():
                    continue
                for pattern in ("*.exe", "*.dll", "*.ocx"):
                    candidates.extend(subdir.rglob(pattern))
        for path in candidates:
            try:
                resolved = path.resolve()
            except OSError:
                continue
            if resolved in seen:
                continue
            if not resolved.is_file():
                continue
            name = resolved.name.lower()
            if name in INTERESTING_NAMES or name.endswith((".ocx", ".bpl")):
                seen.add(resolved)
                found.append(resolved)
            elif any(k in name for k in ("axis", "kbond", "messenger", "ezq", "mss")):
                seen.add(resolved)
                found.append(resolved)
    found.sort(key=lambda p: (str(p).lower(), p.name.lower()))
    return found


def summarize_install(roots: list[Path]) -> list[str]:
    lines: list[str] = ["== install paths =="]
    for root in roots:
        if root.exists():
            kind = "dir" if root.is_dir() else "file"
            lines.append(f"OK  {root} ({kind})")
        else:
            lines.append(f"MISS {root}")
    return lines


def classify_stack(hints: list[FileHint]) -> list[str]:
    lines = ["== stack heuristics (not proof of chat control) =="]
    delphi_files = [h for h in hints if h.delphi_hits]
    mfc_files = [h for h in hints if h.mfc_hits]
    qt_files = [
        h
        for h in hints
        if any(x.startswith("Qt:") or "qt" in x.lower() for x in h.other_hits)
    ]
    electron_files = [
        h for h in hints if any("Electron" in x for x in h.other_hits)
    ]

    if delphi_files:
        lines.append(
            "Delphi/VCL: HINT - Delphi-like strings/imports found "
            f"in {len(delphi_files)} file(s). Not proof TElTree is used for chat."
        )
        for h in delphi_files[:8]:
            lines.append(f"  - {h.path}: {', '.join(h.delphi_hits[:8])}")
    else:
        lines.append(
            "Delphi/VCL: no Delphi/VCL/ElTree string hits in scanned PE blobs"
        )

    if mfc_files:
        lines.append(
            f"MFC: HINT - MFC-like strings/imports in {len(mfc_files)} file(s) "
            "(common for AXIS/MFC clients; does not rule out mixed stacks)."
        )
        for h in mfc_files[:8]:
            lines.append(f"  - {h.path}: {', '.join(h.mfc_hits[:8])}")
    else:
        lines.append("MFC: no MFC string/import hits in scanned PE blobs")

    if qt_files:
        lines.append(f"Qt: HINT - markers in {len(qt_files)} file(s)")
        for h in qt_files[:5]:
            lines.append(f"  - {h.path.name}: {', '.join(h.other_hits[:6])}")
    else:
        lines.append("Qt: no strong hits")

    if electron_files:
        lines.append(f"Electron: HINT - markers in {len(electron_files)} file(s)")
    else:
        lines.append("Electron: no strong hits")

    axis_present = any(h.path.name.lower() == "axis.exe" for h in hints)
    messenger_present = any(
        h.path.name.lower() == "kbondmessenger.exe" for h in hints
    )
    lines.append(
        f"AXIS client binary present: {'yes' if axis_present else 'no'}; "
        f"KBondMessenger.exe present: {'yes' if messenger_present else 'no'}"
    )
    lines.append(
        "Note: trading AXIS (C:\\KBond\\exe) and messenger "
        "(C:\\KBond\\KBondMessenger) are separate trees - chat class must be "
        "checked on the messenger UI, not assumed from axis.exe alone."
    )
    return lines


def dump_pe_details(hints: list[FileHint], verbose: bool) -> list[str]:
    lines = ["== scanned binaries =="]
    for h in hints:
        if h.error:
            lines.append(f"{h.path} ERROR {h.error}")
            continue
        tag = []
        if h.delphi_hits:
            tag.append("delphi?")
        if h.mfc_hits:
            tag.append("mfc?")
        if h.other_hits:
            tag.append("other")
        tag_s = ",".join(tag) if tag else "-"
        lines.append(f"{h.path.name}\t{h.size}\t{tag_s}\t{h.path}")
        if verbose:
            if h.imports:
                lines.append(f"  imports: {', '.join(h.imports[:30])}")
            if h.delphi_hits:
                lines.append(f"  delphi_hits: {', '.join(h.delphi_hits)}")
            if h.mfc_hits:
                lines.append(f"  mfc_hits: {', '.join(h.mfc_hits)}")
            if h.other_hits:
                lines.append(f"  other_hits: {', '.join(h.other_hits)}")
    return lines


def _process_matches() -> list[tuple[int, str, str]]:
    try:
        import psutil
    except ImportError:
        return []
    rows: list[tuple[int, str, str]] = []
    for proc in psutil.process_iter(["pid", "name", "exe"]):
        try:
            name = (proc.info.get("name") or "").lower()
            if name not in PROCESS_NEEDLES_EXACT:
                continue
            exe = proc.info.get("exe") or ""
            rows.append((int(proc.info["pid"]), proc.info.get("name") or "", exe))
        except (psutil.NoSuchProcess, psutil.AccessDenied, TypeError, ValueError):
            continue
    rows.sort(key=lambda r: r[1].lower())
    return rows


def _enum_windows_for_pids(pids: set[int], max_children: int = 80) -> list[str]:
    try:
        import win32gui
        import win32process
    except ImportError:
        return ["pywin32 not available — skip live window class dump"]

    lines: list[str] = ["== live windows (login optional; chat body may be absent) =="]
    tops: list[tuple[int, str, str, int]] = []

    def top_cb(hwnd: int, _: object) -> bool:
        try:
            if not win32gui.IsWindowVisible(hwnd):
                return True
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            if int(pid) not in pids:
                return True
            title = win32gui.GetWindowText(hwnd) or ""
            cls = win32gui.GetClassName(hwnd) or ""
            tops.append((hwnd, cls, title, int(pid)))
        except Exception:
            pass
        return True

    win32gui.EnumWindows(top_cb, None)
    if not tops:
        lines.append("No visible top-level windows for matched PIDs.")
        lines.append(
            "If only background processes run, start KBondMessenger (login "
            "screen is enough for some class hints; chat list needs post-login UI)."
        )
        return lines

    class_counter: Counter[str] = Counter()
    for hwnd, cls, title, pid in tops:
        lines.append(f"TOP hwnd=0x{hwnd:08X} pid={pid} class={cls!r} title={title!r}")
        class_counter[cls] += 1
        children: list[tuple[int, str, str]] = []

        def child_cb(ch: int, __: object) -> bool:
            try:
                ccls = win32gui.GetClassName(ch) or ""
                ctitle = win32gui.GetWindowText(ch) or ""
                children.append((ch, ccls, ctitle))
                class_counter[ccls] += 1
            except Exception:
                pass
            return True

        try:
            win32gui.EnumChildWindows(hwnd, child_cb, None)
        except Exception as exc:
            lines.append(f"  child enum failed: {exc}")
            continue
        for ch, ccls, ctitle in children[:max_children]:
            hint = ""
            low = ccls.lower()
            for key, msg in CLASS_HINTS.items():
                if key in low:
                    hint = f"  << {msg}"
                    break
            shown = ctitle if len(ctitle) <= 60 else ctitle[:57] + "..."
            lines.append(f"  child 0x{ch:08X} class={ccls!r} text={shown!r}{hint}")
        if len(children) > max_children:
            lines.append(f"  ... {len(children) - max_children} more children omitted")

    lines.append("-- class name histogram (top+children) --")
    for cls, count in class_counter.most_common(40):
        note = ""
        low = cls.lower()
        for key, msg in CLASS_HINTS.items():
            if key in low:
                note = f"  ({msg})"
                break
        lines.append(f"  {count:4d}  {cls}{note}")

    if any("teltree" in c.lower() for c in class_counter):
        lines.append(
            "VERDICT class: TElTree appears in live UI - ElTree usage for some "
            "control is strongly suggested (confirm it is the chat message list)."
        )
    else:
        lines.append(
            "VERDICT class: TElTree not seen among enumerated classes. "
            "Do not assume ElTree API yet (may simply be pre-login / no chat UI)."
        )
    return lines


def conclusions() -> list[str]:
    return [
        "== how to read this ==",
        "Confirmed without login: install layout, AXIS/Messenger binaries, PE string/import hints.",
        "Not confirmed without chat UI: message-list HWND class, UIA text exposure, need for ElTree API.",
        "Next after login: aim Spy++/Inspect at the chat body; only if class is TElTree treat ElTree as likely.",
        "Prefer messenger process windows over axis.exe when diagnosing chat reading.",
    ]


def _safe_print(text: str) -> None:
    encoding = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        sys.stdout.write(text + "\n")
    except UnicodeEncodeError:
        sys.stdout.buffer.write((text + "\n").encode(encoding, errors="replace"))
        sys.stdout.buffer.flush()


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="K-Bond stack hints without requiring login (PE + optional live windows)."
    )
    p.add_argument(
        "--root",
        action="append",
        type=Path,
        default=[],
        help="Extra install root (repeatable). Defaults include C:\\KBond trees.",
    )
    p.add_argument(
        "--no-default-roots",
        action="store_true",
        help="Do not use built-in C:\\KBond paths",
    )
    p.add_argument(
        "--max-bytes",
        type=int,
        default=12_000_000,
        help="Max bytes to read per PE for string/import scan",
    )
    p.add_argument(
        "--skip-live",
        action="store_true",
        help="Skip process/window enumeration",
    )
    p.add_argument(
        "--verbose",
        action="store_true",
        help="Print imports and hit lists per file",
    )
    return p.parse_args(argv)


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    roots: list[Path] = []
    if not args.no_default_roots:
        roots.extend(DEFAULT_ROOTS)
    roots.extend(args.root)
    # de-dupe preserve order
    uniq_roots: list[Path] = []
    seen_r: set[Path] = set()
    for r in roots:
        key = r
        try:
            key = r.resolve()
        except OSError:
            pass
        if key in seen_r:
            continue
        seen_r.add(key)
        uniq_roots.append(r)

    out: list[str] = []
    out.extend(summarize_install(uniq_roots))
    binaries = discover_binaries(uniq_roots)
    out.append(f"== discovered {len(binaries)} interesting binaries ==")
    hints = [analyze_pe(path, max_bytes=args.max_bytes) for path in binaries]
    out.extend(dump_pe_details(hints, verbose=args.verbose))
    out.extend(classify_stack(hints))

    if not args.skip_live:
        procs = _process_matches()
        out.append("== matching processes ==")
        if not procs:
            out.append("None (psutil missing or KBond/axis not running).")
        else:
            pids: set[int] = set()
            for pid, name, exe in procs:
                out.append(f"pid={pid} name={name} exe={exe or '(path unavailable)'}")
                pids.add(pid)
            out.extend(_enum_windows_for_pids(pids))

    out.extend(conclusions())
    _safe_print("\n".join(out))
    return 0


if __name__ == "__main__":
    sys.exit(main())
