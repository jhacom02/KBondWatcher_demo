from __future__ import annotations

import sys
from datetime import date, datetime
from pathlib import Path
from typing import Optional

from .deploy_mode import is_dev, is_pilot
from .license import LicenseError


def demo_expiry_candidates() -> list[Path]:
    """Prefer file next to frozen exe / interpreter entry folder."""
    out: list[Path] = []
    exe = Path(sys.executable).resolve()
    out.append(exe.parent / "demo_expiry.txt")
    try:
        main = Path(sys.modules["__main__"].__file__ or "").resolve()
        if main.is_file():
            out.append(main.parent / "demo_expiry.txt")
    except Exception:
        pass
    # Dedupe while preserving order
    seen: set[Path] = set()
    unique: list[Path] = []
    for p in out:
        if p in seen:
            continue
        seen.add(p)
        unique.append(p)
    return unique


def find_demo_expiry_file() -> Optional[Path]:
    for path in demo_expiry_candidates():
        if path.is_file():
            return path
    return None


def parse_demo_expiry(text: str) -> date:
    raw = (text or "").strip().splitlines()[0].strip() if text else ""
    if not raw:
        raise LicenseError("demo_expiry.txt is empty")
    try:
        return datetime.strptime(raw, "%Y-%m-%d").date()
    except ValueError as exc:
        raise LicenseError(f"demo_expiry.txt invalid date: {raw!r}") from exc


def check_demo_expiry(*, today: Optional[date] = None) -> None:
    """
    Fail-closed in pilot when missing/expired.
    Dev skips when file is absent.
    """
    path = find_demo_expiry_file()
    if path is None:
        if is_pilot():
            raise LicenseError("demo_expiry.txt missing (pilot fail-closed)")
        if is_dev():
            return
        raise LicenseError("demo_expiry.txt missing")
    expiry = parse_demo_expiry(path.read_text(encoding="utf-8"))
    current = today or date.today()
    if current > expiry:
        raise LicenseError(f"demo expired on {expiry.isoformat()} (today={current.isoformat()})")
