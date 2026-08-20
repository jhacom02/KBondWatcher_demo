from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.models import Quote  # noqa: E402
from core.trigger import (  # noqa: E402
    evaluate,
    flip_side_token,
    format_message,
    format_looking_for_label,
    looking_for_from_qty,
    pnl_outside_sanity_band,
)


def test_format_looking_for_label() -> None:
    assert format_looking_for_label("25-11", "BID") == "25-11 / BID"
    assert format_looking_for_label("25-10", "OFFER") == "25-10 / OFFER"


def test_looking_for_from_qty() -> None:
    assert looking_for_from_qty(-100) == ("BID", "BUY")
    assert looking_for_from_qty(100) == ("OFFER", "SELL")
    assert looking_for_from_qty(-80) == ("BID", "BUY")
    assert looking_for_from_qty(80) == ("OFFER", "SELL")
    with pytest.raises(ValueError):
        looking_for_from_qty(0)
    with pytest.raises(ValueError):
        looking_for_from_qty(80.5)


def test_evaluate_le() -> None:
    quote = Quote("25-11", "25-11 23+", "23+", 4.23, "BUY")
    assert evaluate(quote, 100_000, 100_000, "<=").triggered is True
    assert evaluate(quote, 200_000, 100_000, "<=").triggered is False
    assert evaluate(quote, -2_000_000, 100_000, "<=").triggered is True


def test_evaluate_le_negative_threshold() -> None:
    quote = Quote("25-11", "25-11 23+", "23+", 4.23, "BUY")
    assert evaluate(quote, -2_000_000, -2_000_000, "<=").triggered is True
    assert evaluate(quote, -1_000_000, -2_000_000, "<=").triggered is False


def test_pnl_outside_sanity_band() -> None:
    assert pnl_outside_sanity_band(-189049, -80, 5_000_000) is False
    assert pnl_outside_sanity_band(-5_000_080, -80, 5_000_000) is False
    assert pnl_outside_sanity_band(-5_000_081, -80, 5_000_000) is True
    assert pnl_outside_sanity_band(-2146826273, -80, 5_000_000) is True


def test_evaluate_ge() -> None:
    quote = Quote("25-11", "25-11 23-", "23-", 4.23, "SELL")
    assert evaluate(quote, 1_500_000, 1_000_000, ">=").triggered is True
    assert evaluate(quote, 1_000_000, 1_000_000, ">=").triggered is True
    assert evaluate(quote, 500_000, 1_000_000, ">=").triggered is False


def test_flip_side_token() -> None:
    assert flip_side_token("715+") == "715-"
    assert flip_side_token("715-") == "715+"
    assert flip_side_token("23사자") == "23팔자"
    assert flip_side_token("23팔자") == "23사자"


def test_format_message_confirm_token() -> None:
    quote = Quote("25-10", "25-10 715+", "715+", 3.715, "BUY")
    text = format_message("{instrument} {confirm_token} ㅎㅈ", quote, 1500000)
    assert text == "25-10 715- ㅎㅈ"


def test_format_message_appends_qty_when_not_100() -> None:
    quote = Quote("25-10", "25-10 695 + 80억", "695 +", 3.695, "BUY", quantity=80)
    text = format_message("{instrument} {confirm_token} ㅎㅈ", quote, 1500000)
    assert text == "25-10 695 - 80억 ㅎㅈ"
