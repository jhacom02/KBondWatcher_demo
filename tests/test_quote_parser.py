"""Unit tests for quote_parser."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from quote_parser import parse_quote_line  # noqa: E402


@pytest.mark.parametrize(
    "line,expected_yield,expected_side",
    [
        ("25-11 23+", 4.23, "BUY"),
        ("25-11 23-", 4.23, "SELL"),
        ("25-11 235+", 4.235, "BUY"),
        ("25-11 235-", 4.235, "SELL"),
        ("홍길동 (16:20:30) : 25-11 23+ 100억 (...)", 4.23, "BUY"),
        ("25-11 23 사자", 4.23, "BUY"),
        ("25-11 23 팔자", 4.23, "SELL"),
    ],
)
def test_parse_basic_quotes(line: str, expected_yield: float, expected_side: str) -> None:
    quote = parse_quote_line(line, target="25-11", yield_prefix=4)
    assert quote is not None
    assert quote.instrument == "25-11"
    assert quote.yield_value == pytest.approx(expected_yield)
    assert quote.side == expected_side


@pytest.mark.parametrize(
    "line",
    [
        "125-11 23+",
        "25-110 23+",
    ],
)
def test_partial_match_prevention(line: str) -> None:
    quote = parse_quote_line(line, target="25-11", yield_prefix=4)
    assert quote is None


def test_required_side_filter() -> None:
    buy = parse_quote_line("25-11 23+", target="25-11", yield_prefix=4, required_side="BUY")
    assert buy is not None
    sell_filtered = parse_quote_line(
        "25-11 23+", target="25-11", yield_prefix=4, required_side="SELL"
    )
    assert sell_filtered is None


def test_no_quote_token() -> None:
    assert parse_quote_line("25-11 hello", target="25-11", yield_prefix=4) is None
