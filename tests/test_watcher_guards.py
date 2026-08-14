from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel import InstrumentSlot
from main import (
    LINE_LOG_MAX_PER_POLL,
    _truncate_log_text,
    collect_batch_matches,
    log_new_source_lines,
    watch_identity,
)
from source import SourceReaderError, message_fingerprint
from source.common import source_line


def _slot() -> InstrumentSlot:
    return InstrumentSlot(
        instrument="25-10",
        row=41,
        looking_for="BID",
        required_side="BUY",
        qty_abs=100,
        yield_prefix=3.0,
        input_cell="D41",
        qty_cell="E41",
        pnl_cell="F44",
    )


def test_collect_batch_matches_two_quotes_raises() -> None:
    slot = _slot()
    lines = [
        source_line("25-10 695 +", "(17:48:01) : 25-10 695 +"),
        source_line("25-10 696 +", "(17:48:02) : 25-10 696 +"),
    ]
    with pytest.raises(SourceReaderError, match="ambiguous quotes in one poll: 2"):
        collect_batch_matches(lines, [slot], set())


def test_collect_batch_matches_one_quote() -> None:
    slot = _slot()
    lines = [
        source_line("noise"),
        source_line("25-10 695 +", "(17:48:01) : 25-10 695 +"),
    ]
    matches = collect_batch_matches(lines, [slot], set())
    assert len(matches) == 1
    assert matches[0][1].raw_token == "695 +"


def test_collect_batch_matches_skips_fingerprint() -> None:
    slot = _slot()
    line = source_line("25-10 695 +", "(17:48:01) : 25-10 695 +")
    fp = message_fingerprint(line.watermark_key)
    matches = collect_batch_matches([line], [slot], {fp})
    assert matches == []


def test_collect_batch_matches_require_clock_skips() -> None:
    slot = _slot()
    lines = [source_line("25-10 695 +")]
    matches = collect_batch_matches(lines, [slot], set(), require_clock=True)
    assert matches == []


def test_watch_identity_tuple() -> None:
    slot = _slot()
    assert watch_identity(slot, 100_000.0) == ("25-10", "BID", 100, 100_000.0)


def test_truncate_log_text() -> None:
    assert _truncate_log_text("short") == "short"
    long = "a" * 200
    out = _truncate_log_text(long)
    assert len(out) == 160
    assert out.endswith("…")


class _ListLog:
    def __init__(self) -> None:
        self.messages: list[str] = []

    def info(self, fmt: str, *args) -> None:
        self.messages.append(fmt % args)


def test_log_new_source_lines_cap_and_mode3_key() -> None:
    log = _ListLog()
    lines = [source_line(f"body-{i}", f"(12:00:{i:02d}) : body-{i}") for i in range(25)]
    log_new_source_lines(
        log,
        mode=3,
        looking_for="BID",
        threshold=0.0,
        lines=lines,
    )
    assert sum(1 for m in log.messages if m.startswith("LINE |")) == LINE_LOG_MAX_PER_POLL
    assert any(m.startswith("LINE_OMITTED | +5") for m in log.messages)
    assert log.messages[0].startswith(
        "LINE | mode=3 | looking_for=BID | threshold=0 | raw_line="
    )
    assert "(12:00:00) : body-0" in log.messages[0]

