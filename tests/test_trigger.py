from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from models import Quote  # noqa: E402
from trigger import evaluate, format_message  # noqa: E402


def test_evaluate_trigger() -> None:
    quote = Quote("25-11", "25-11 23+", "23+", 4.23, "BUY")
    ok = evaluate(quote, 1_500_000, 1_000_000)
    assert ok.triggered is True
    no = evaluate(quote, 500_000, 1_000_000)
    assert no.triggered is False


def test_format_message() -> None:
    quote = Quote("25-10", "25-10 23+", "23+", 3.23, "BUY")
    text = format_message("{instrument} {raw_token} ㅎㅈ", quote, 1500000)
    assert text == "25-10 23+ ㅎㅈ"
