"""Unit tests for Excel bridge pure helpers (no COM required)."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel_bridge import (  # noqa: E402
    ExcelBridgeError,
    format_status,
    to_float,
)
from models import AppStatus  # noqa: E402


def test_to_float_number() -> None:
    assert to_float(1532000) == 1532000.0
    assert to_float(4.23) == pytest.approx(4.23)


def test_to_float_string_with_comma() -> None:
    assert to_float("1,532,000") == 1532000.0


def test_to_float_rejects_empty() -> None:
    with pytest.raises(ExcelBridgeError):
        to_float(None)
    with pytest.raises(ExcelBridgeError):
        to_float("")
    with pytest.raises(ExcelBridgeError):
        to_float("  ")


def test_format_status() -> None:
    assert format_status(AppStatus.SENT) == "SENT"
    assert format_status("WATCHING") == "WATCHING"
