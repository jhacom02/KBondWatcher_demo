from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from source.quote_parser import parse_quote_line


@pytest.mark.parametrize(
    "line,target,prefix,expected_yield,expected_side",
    [
        ("25-10 735+", "25-10", 3, 3.735, "BUY"),
        ("25-10 735 +", "25-10", 3, 3.735, "BUY"),
        ("25-10 735사자", "25-10", 3, 3.735, "BUY"),
        ("25-10 735 사자", "25-10", 3, 3.735, "BUY"),
        ("25-11 23+", "25-11", 4, 4.23, "BUY"),
        ("25-11 23-", "25-11", 4, 4.23, "SELL"),
        ("25-11 235+", "25-11", 4, 4.235, "BUY"),
        ("25-11 235-", "25-11", 4, 4.235, "SELL"),
        ("25-11 23 사자", "25-11", 4, 4.23, "BUY"),
        ("25-11 23 팔자", "25-11", 4, 4.23, "SELL"),
        ("홍길동 (16:20:30) : 25-11 23+", "25-11", 4, 4.23, "BUY"),
    ],
)
def test_accept(
    line: str,
    target: str,
    prefix: float,
    expected_yield: float,
    expected_side: str,
) -> None:
    quote = parse_quote_line(line, target=target, yield_prefix=prefix)
    assert quote is not None
    assert quote.instrument == target
    assert quote.yield_value == pytest.approx(expected_yield)
    assert quote.side == expected_side


@pytest.mark.parametrize(
    "line,target,prefix",
    [
        ("25-10 73 팔자 40억", "25-10", 3),
        ("25-10 73 - 40억", "25-10", 3),
        ("25-10 73 팔자 자투리", "25-10", 3),
        ("25-10 73- 자투리", "25-10", 3),
        ("25-10 사고 25-11 파는 교체", "25-10", 3),
        ("25-10 매수있나요?", "25-10", 3),
        ("25-10 사자호가 찾습니다", "25-10", 3),
        ("25-10 735+ ㅎㅈ", "25-10", 3),
        ("25-10 73/735-", "25-10", 3),
        ("25-10 735- 동", "25-10", 3),
        ("25-10 735- 선", "25-10", 3),
        ("125-11 23+", "25-11", 4),
        ("25-110 23+", "25-11", 4),
        ("25-11 hello", "25-11", 4),
        ("홍길동 (16:20:30) : 25-11 23+ 100억 (...)", "25-11", 4),
    ],
)
def test_reject(line: str, target: str, prefix: float) -> None:
    assert parse_quote_line(line, target=target, yield_prefix=prefix) is None


def test_required_side_filter() -> None:
    buy = parse_quote_line("25-11 23+", target="25-11", yield_prefix=4, required_side="BUY")
    assert buy is not None
    assert (
        parse_quote_line("25-11 23+", target="25-11", yield_prefix=4, required_side="SELL")
        is None
    )
