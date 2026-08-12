from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from models import Quote  # noqa: E402
from trigger import (  # noqa: E402
    evaluate,
    flip_side_token,
    format_message,
    looking_for_from_qty,
)


def test_looking_for_from_qty() -> None:
    assert looking_for_from_qty(-100) == ("BID", "BUY")
    assert looking_for_from_qty(100) == ("OFFER", "SELL")
    with pytest.raises(ValueError):
        looking_for_from_qty(0)


def test_evaluate_bid() -> None:
    quote = Quote("25-11", "25-11 23+", "23+", 4.23, "BUY")
    ok = evaluate(quote, 1_500_000, 1_000_000, "BID")
    assert ok.triggered is True
    no = evaluate(quote, 500_000, 1_000_000, "BID")
    assert no.triggered is False


def test_evaluate_offer() -> None:
    quote = Quote("25-11", "25-11 23-", "23-", 4.23, "SELL")
    ok = evaluate(quote, -1_500_000, 1_000_000, "OFFER")
    assert ok.triggered is True
    no = evaluate(quote, -500_000, 1_000_000, "OFFER")
    assert no.triggered is False


def test_flip_side_token() -> None:
    assert flip_side_token("715+") == "715-"
    assert flip_side_token("715-") == "715+"
    assert flip_side_token("23사자") == "23팔자"
    assert flip_side_token("23팔자") == "23사자"


def test_format_message_confirm_token() -> None:
    quote = Quote("25-10", "25-10 715+", "715+", 3.715, "BUY")
    text = format_message("{instrument} {confirm_token} ㅎㅈ", quote, 1500000)
    assert text == "25-10 715- ㅎㅈ"
