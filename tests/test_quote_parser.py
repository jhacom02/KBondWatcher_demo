from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from source.quote_parser import parse_quote_line


@pytest.mark.parametrize(
    "line,target,prefix,expected_yield,expected_side,expected_token",
    [
        ("25-10 735+", "25-10", 3, 3.735, "BUY", "735+"),
        ("25-10 735 +", "25-10", 3, 3.735, "BUY", "735 +"),
        ("25-10 735사자", "25-10", 3, 3.735, "BUY", "735사자"),
        ("25-10 735 사자", "25-10", 3, 3.735, "BUY", "735 사자"),
        ("25-11 23+", "25-11", 4, 4.23, "BUY", "23+"),
        ("25-11 23-", "25-11", 4, 4.23, "SELL", "23-"),
        ("25-11 235+", "25-11", 4, 4.235, "BUY", "235+"),
        ("25-11 235-", "25-11", 4, 4.235, "SELL", "235-"),
        ("25-11 23 사자", "25-11", 4, 4.23, "BUY", "23 사자"),
        ("25-11 23 팔자", "25-11", 4, 4.23, "SELL", "23 팔자"),
        ("홍길동 (16:20:30) : 25-11 23+", "25-11", 4, 4.23, "BUY", "23+"),
        (
            "신영환 (11:06:38) : 22-14  14- (**증권 채권금융 368-****)",
            "22-14",
            3,
            3.14,
            "SELL",
            "14-",
        ),
        ("25-10 735+ *증권", "25-10", 3, 3.735, "BUY", "735+"),
        ("25-10 695 + 100", "25-10", 3, 3.695, "BUY", "695 +"),
        ("25-10 695 + 100억", "25-10", 3, 3.695, "BUY", "695 +"),
        (
            "홍길동 (16:20:30) : 25-11 23+ 100억 (...)",
            "25-11",
            4,
            4.23,
            "BUY",
            "23+",
        ),
    ],
)
def test_accept(
    line: str,
    target: str,
    prefix: float,
    expected_yield: float,
    expected_side: str,
    expected_token: str,
) -> None:
    quote = parse_quote_line(line, target=target, yield_prefix=prefix)
    assert quote is not None
    assert quote.instrument == target
    assert quote.yield_value == pytest.approx(expected_yield)
    assert quote.side == expected_side
    assert quote.raw_token == expected_token
    assert quote.raw_line == line.strip()


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
        ("25-10 695 + 80억", "25-10", 3),
        ("25-10 695 + 80", "25-10", 3),
        ("25-10 695 + 100억 있나요", "25-10", 3),
        ("25-10 695 + 100 있나요", "25-10", 3),
        ("25-10 695 + 100 80", "25-10", 3),
        ("25-10 695 + 있으신가요", "25-10", 3),
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


def test_required_qty_100_accepts_implied_and_explicit() -> None:
    implied = parse_quote_line(
        "25-10 695 +", target="25-10", yield_prefix=3, required_qty=100
    )
    assert implied is not None
    assert implied.quantity == 100
    explicit = parse_quote_line(
        "25-10 695 + 100억", target="25-10", yield_prefix=3, required_qty=100
    )
    assert explicit is not None
    assert explicit.quantity == 100
    assert explicit.raw_token == "695 +"


def test_required_qty_80_requires_explicit() -> None:
    assert (
        parse_quote_line(
            "25-10 695 +", target="25-10", yield_prefix=3, required_qty=80
        )
        is None
    )
    assert (
        parse_quote_line(
            "25-10 695 + 100억", target="25-10", yield_prefix=3, required_qty=80
        )
        is None
    )
    hit = parse_quote_line(
        "25-10 695 + 80억", target="25-10", yield_prefix=3, required_qty=80
    )
    assert hit is not None
    assert hit.quantity == 80
    hit_num = parse_quote_line(
        "25-10 695 + 80", target="25-10", yield_prefix=3, required_qty=80
    )
    assert hit_num is not None
    assert hit_num.quantity == 80
