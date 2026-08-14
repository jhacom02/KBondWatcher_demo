from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from source.reader_uia import attach_preceding_time


def test_attach_preceding_time_binds_immediate_time_line() -> None:
    items = attach_preceding_time(
        [
            "권** (17:48:01) :",
            "26-3 005 01 + 20",
            "(**증권 종합금융팀 02-****)",
        ]
    )
    assert items[0].text == "권** (17:48:01) :"
    assert items[0].watermark_key == "권** (17:48:01) :"
    assert items[1].text == "26-3 005 01 + 20"
    assert items[1].watermark_key == "(17:48:01) : 26-3 005 01 + 20"
    assert items[2].text == "(**증권 종합금융팀 02-****)"
    assert items[2].watermark_key == "(**증권 종합금융팀 02-****)"


def test_attach_preceding_time_bare_clock() -> None:
    items = attach_preceding_time(["(17:48:01) :", "25-10 695 +"])
    assert items[1].text == "25-10 695 +"
    assert items[1].watermark_key == "(17:48:01) : 25-10 695 +"


def test_attach_preceding_time_no_time_keeps_quote() -> None:
    items = attach_preceding_time(["26-3 005 01 + 20"])
    assert items[0].text == "26-3 005 01 + 20"
    assert items[0].watermark_key == "26-3 005 01 + 20"


def test_attach_preceding_time_does_not_split_combined_line() -> None:
    full = "권** (17:48:01) : 26-3 005 01 + 20"
    items = attach_preceding_time([full])
    assert len(items) == 1
    assert items[0].text == full
    assert items[0].watermark_key == full


def test_attach_preceding_time_same_quote_two_clocks() -> None:
    items = attach_preceding_time(
        [
            "(17:48:01) :",
            "26-3 005 01 + 20",
            "(17:48:02) :",
            "26-3 005 01 + 20",
        ]
    )
    quotes = [item for item in items if item.text == "26-3 005 01 + 20"]
    assert [item.watermark_key for item in quotes] == [
        "(17:48:01) : 26-3 005 01 + 20",
        "(17:48:02) : 26-3 005 01 + 20",
    ]
