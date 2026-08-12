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
    normalize_instrument,
    prefix_from_prev_yield,
    to_float,
    workbook_matches_open,
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


def test_workbook_matches_absolute_full_name() -> None:
    cfg = r"C:\Users\Daily\daily.xlsm"
    assert workbook_matches_open(cfg, "daily.xlsm", r"C:\Users\Daily\daily.xlsm")
    assert workbook_matches_open(cfg, "other.xlsm", r"C:\Users\Daily\other.xlsm") is False


def test_workbook_matches_by_file_name() -> None:
    cfg = r"C:\Users\Daily\daily.xlsm"
    assert workbook_matches_open(cfg, "daily.xlsm", r"D:\elsewhere\daily.xlsm")


def test_normalize_instrument() -> None:
    assert normalize_instrument("국고 25-5") == "25-5"
    assert normalize_instrument("국고25-10") == "25-10"
    assert normalize_instrument("  국고  25-11 ") == "25-11"
    assert normalize_instrument("") == ""


def test_prefix_from_prev_yield() -> None:
    assert prefix_from_prev_yield(3.215) == 3.0
    assert prefix_from_prev_yield(4.180) == 4.0
    assert prefix_from_prev_yield(-3.99) == 3.0
