from __future__ import annotations

import csv
import statistics
from datetime import datetime
from pathlib import Path
from typing import Iterable, Optional

from .logger import get_logger

COLUMNS = (
    "ts",
    "total_ms",
    "excel_ms",
    "send_ms",
    "mode",
    "looking_for",
    "raw_line",
    "sent_message",
)

MS_FIELDS = ("total_ms", "excel_ms", "send_ms")


def sent_perf_path(log_path: Path) -> Path:
    return Path(log_path).parent / "sent_perf.csv"


def _ms(value: float) -> str:
    return f"{value:.1f}"


def append_sent(
    path: Path,
    *,
    mode: int,
    looking_for: str,
    raw_line: str,
    sent_message: str,
    total_ms: float,
    excel_ms: float,
    send_ms: float,
    ts: Optional[datetime] = None,
) -> None:
    path = Path(path)
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        new_file = not path.is_file() or path.stat().st_size == 0
        with path.open("a", encoding="utf-8", newline="") as fh:
            writer = csv.DictWriter(fh, fieldnames=COLUMNS, lineterminator="\n")
            if new_file:
                writer.writeheader()
            writer.writerow(
                {
                    "ts": (ts or datetime.now()).isoformat(timespec="seconds"),
                    "total_ms": _ms(total_ms),
                    "excel_ms": _ms(excel_ms),
                    "send_ms": _ms(send_ms),
                    "mode": str(mode),
                    "looking_for": looking_for,
                    "raw_line": raw_line,
                    "sent_message": sent_message,
                }
            )
    except OSError as exc:
        get_logger().error("Failed to append sent perf log %s: %s", path, exc)


def _floats(rows: Iterable[dict[str, str]], key: str) -> list[float]:
    out: list[float] = []
    for row in rows:
        raw = (row.get(key) or "").strip()
        if not raw:
            continue
        out.append(float(raw))
    return out


def _stats_line(label: str, values: list[float]) -> str:
    if not values:
        return f"{label} count=0"
    return (
        f"{label} count={len(values)} "
        f"mean={statistics.mean(values):.1f} "
        f"median={statistics.median(values):.1f}"
    )


def summarize(path: Path) -> str:
    path = Path(path)
    if not path.is_file() or path.stat().st_size == 0:
        return f"no sent perf log: {path}"
    with path.open("r", encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    if not rows:
        return f"no sent perf rows: {path}"
    lines = [f"file={path}"]
    for field in MS_FIELDS:
        lines.append(_stats_line(field, _floats(rows, field)))
    modes = sorted({(row.get("mode") or "").strip() for row in rows if (row.get("mode") or "").strip()})
    for mode in modes:
        subset = [row for row in rows if (row.get("mode") or "").strip() == mode]
        lines.append(_stats_line(f"mode={mode} total_ms", _floats(subset, "total_ms")))
    return "\n".join(lines)
