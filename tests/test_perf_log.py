from __future__ import annotations

import sys
from datetime import datetime
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.perf_log import append_sent, sent_perf_path, summarize


def test_sent_perf_path_next_to_log() -> None:
    assert sent_perf_path(Path("logs/watcher.log")) == Path("logs") / "sent_perf.csv"


def test_append_sent_and_summarize(tmp_path: Path) -> None:
    path = tmp_path / "sent_perf.csv"
    stamp = datetime(2026, 8, 14, 16, 0, 0)
    append_sent(
        path,
        mode=2,
        looking_for="25-10 / BID",
        raw_line="25-10 695 +",
        sent_message="25-10 695 - ㅎㅈ",
        total_ms=10.0,
        excel_ms=4.0,
        send_ms=5.0,
        ts=stamp,
    )
    append_sent(
        path,
        mode=2,
        looking_for="25-10 / BID",
        raw_line="25-10 696 +",
        sent_message="25-10 696 - ㅎㅈ",
        total_ms=20.0,
        excel_ms=6.0,
        send_ms=12.0,
        ts=stamp,
    )
    text = path.read_text(encoding="utf-8")
    assert text.startswith(
        "ts,total_ms,excel_ms,send_ms,mode,looking_for,raw_line,sent_message\n"
    )
    assert "25-10 695 +" in text
    assert "25-10 695 - ㅎㅈ" in text
    summary = summarize(path)
    assert "count=2" in summary
    assert "total_ms count=2 mean=15.0 median=15.0" in summary
    assert "excel_ms count=2 mean=5.0 median=5.0" in summary
    assert "mode=2 total_ms count=2 mean=15.0 median=15.0" in summary


def test_summarize_missing_file(tmp_path: Path) -> None:
    path = tmp_path / "missing.csv"
    assert summarize(path).startswith("no sent perf log:")
