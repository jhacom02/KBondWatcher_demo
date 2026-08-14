from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from excel.bridge import (  # noqa: E402
    ExcelBridge,
    ExcelBridgeError,
    StopRequested,
    bind_slot_cells,
    format_status,
    is_excel_busy,
    normalize_instrument,
    parse_watch_row,
    prefix_from_prev_yield,
    to_float,
    workbook_matches_open,
)
from core.models import AppStatus  # noqa: E402


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


class _FakeComError(Exception):
    pass


def _dummy_bridge() -> ExcelBridge:
    return ExcelBridge(
        workbook_name="sample.xlsm",
        sheet_name="트레이딩",
        status_cell="F2",
        looking_for_cell="G2",
        last_quote_cell="H2",
        last_pnl_cell="I2",
        last_action_cell="J2",
        slot_rows=[19],
        rows_10y=[19],
        rows_3y=[41],
        prefix_3y_cell="B5",
        prefix_10y_cell="B6",
        instrument_col="A",
        qty_col="E",
        input_col="D",
        pnl_col="F",
        pnl_row_offset=3,
    )


def test_is_excel_busy_signed_hresult() -> None:
    exc = _FakeComError(-2147417846, "The message filter indicated that the application is busy.")
    assert is_excel_busy(exc) is True


def test_is_excel_busy_unsigned_hresult() -> None:
    exc = _FakeComError(-2147417846 & 0xFFFFFFFF, "busy")
    assert is_excel_busy(exc) is True


def test_is_excel_busy_call_rejected() -> None:
    exc = _FakeComError(-2147418111, "call was rejected by callee")
    assert is_excel_busy(exc) is True


def test_is_excel_busy_other_error() -> None:
    assert is_excel_busy(RuntimeError("nope")) is False
    assert is_excel_busy(_FakeComError("not-an-hresult")) is False


def test_call_excel_retries_then_succeeds(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("excel.bridge.COM_BUSY_RETRY_DELAY_SECONDS", 0)
    bridge = _dummy_bridge()
    hits = {"n": 0}

    def _flaky() -> str:
        hits["n"] += 1
        if hits["n"] < 3:
            raise _FakeComError(-2147417846, "busy")
        return "ok"

    assert bridge._call_excel(_flaky) == "ok"
    assert hits["n"] == 3


def test_call_excel_raises_when_retries_exhausted(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("excel.bridge.COM_BUSY_RETRY_DELAY_SECONDS", 0)
    monkeypatch.setattr("excel.bridge.COM_BUSY_RETRY_COUNT", 3)
    bridge = _dummy_bridge()

    def _always_busy() -> None:
        raise _FakeComError(-2147417846, "busy")

    with pytest.raises(ExcelBridgeError, match="Excel busy after"):
        bridge._call_excel(_always_busy)


def test_call_excel_stop_flag_raises() -> None:
    bridge = _dummy_bridge()
    bridge.set_stop_check(lambda: True)
    with pytest.raises(StopRequested):
        bridge._call_excel(lambda: "ok")


def test_pid_file_path() -> None:
    from main import pid_file_path

    assert pid_file_path(Path(r"C:\temp\kbond_watcher.stop")) == Path(
        r"C:\temp\kbond_watcher.pid"
    )


_SLOT_ROWS = (19, 25, 41, 46, 56)
_INSTRUMENTS = {
    19: "국고 25-5",
    25: "국고 25-11",
    41: "국고 25-10",
    46: "국고 25-04",
    56: "국고 25-08",
}


def test_bind_slot_cells_row_41() -> None:
    input_cell, qty_cell, pnl_cell = bind_slot_cells(41, "D", "E", "F", 3)
    assert input_cell == "D41"
    assert qty_cell == "E41"
    assert pnl_cell == "F44"


def test_parse_watch_row_formula_a41() -> None:
    assert (
        parse_watch_row("=A41", "국고 25-10", _SLOT_ROWS, _INSTRUMENTS, "트레이딩")
        == 41
    )
    assert parse_watch_row("=$A$19", None, _SLOT_ROWS, _INSTRUMENTS) == 19
    assert (
        parse_watch_row(
            "=트레이딩!A25", "국고 25-11", _SLOT_ROWS, _INSTRUMENTS, "트레이딩"
        )
        == 25
    )
    assert (
        parse_watch_row(
            "='트레이딩'!A46", None, _SLOT_ROWS, _INSTRUMENTS, "트레이딩"
        )
        == 46
    )


def test_parse_watch_row_plain_instrument() -> None:
    assert (
        parse_watch_row("국고 25-10", "국고 25-10", _SLOT_ROWS, _INSTRUMENTS)
        == 41
    )


def test_parse_watch_row_rejects_bad_formula() -> None:
    with pytest.raises(ExcelBridgeError):
        parse_watch_row("=B41", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        parse_watch_row("=A41+0", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        parse_watch_row("=A99", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        parse_watch_row("=Other!A41", None, _SLOT_ROWS, _INSTRUMENTS, "트레이딩")
    with pytest.raises(ExcelBridgeError):
        parse_watch_row("", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        parse_watch_row("국고 99-99", "국고 99-99", _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        parse_watch_row(
            "25-10",
            "25-10",
            _SLOT_ROWS,
            {19: "국고 25-10", 41: "국고 25-10"},
        )
