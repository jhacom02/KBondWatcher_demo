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
    ExcelDisconnected,
    StopRequested,
    bind_slot_cells,
    is_excel_busy,
    is_excel_gone,
    normalize_instrument,
    parse_watch_row,
    pick_matching_workbook,
    prefix_from_prev_yield,
    rot_names_matching_workbook,
    to_float,
    workbook_bind_paths,
    workbook_identity,
    workbook_matches_open,
)


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


def test_to_float_rejects_cverr() -> None:
    with pytest.raises(ExcelBridgeError, match="#VALUE!"):
        to_float(-2146826273)
    with pytest.raises(ExcelBridgeError, match="#VALUE!"):
        to_float("#VALUE!")
    with pytest.raises(ExcelBridgeError, match="#N/A"):
        to_float("#N/A")


def test_write_yield_read_pnl_retries_cverr(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("excel.bridge.CALC_POLL_INTERVAL_SECONDS", 0)
    bridge = _dummy_bridge()
    values = [-2146826273, -189049]

    class _App:
        CalculationState = 0

    class _Cell:
        def __init__(self) -> None:
            self.Value: object = None

    pnl_cell = _Cell()

    class _Ws:
        def Range(self, addr: str) -> _Cell:
            if addr == "F44":
                pnl_cell.Value = values.pop(0) if values else -189049
                return pnl_cell
            return _Cell()

    bridge._app = _App()
    bridge._ws = _Ws()
    monkeypatch.setattr(bridge, "_call_excel", lambda fn, check_stop=True: fn())
    monkeypatch.setattr(bridge, "write_yield", lambda *a, **k: None)
    monkeypatch.setattr(bridge, "_ensure", lambda: None)
    assert bridge.write_yield_read_pnl("D41", "F44", 3.7) == -189049.0


def test_workbook_identity_fullname_exception() -> None:
    class _Wb:
        Name = "sample.xlsm"

        @property
        def FullName(self) -> str:
            raise RuntimeError("COM FullName failed")

    with pytest.raises(ExcelBridgeError, match="FullName"):
        workbook_identity(_Wb())


def test_write_yield_read_pnl_timeout_on_cverr(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("excel.bridge.CALC_POLL_INTERVAL_SECONDS", 0)
    bridge = _dummy_bridge()
    bridge.calc_wait_timeout_seconds = 0

    class _App:
        CalculationState = 0

    class _Cell:
        Value = -2146826273

    class _Ws:
        def Range(self, addr: str) -> _Cell:
            return _Cell()

    bridge._app = _App()
    bridge._ws = _Ws()
    monkeypatch.setattr(bridge, "_call_excel", lambda fn, check_stop=True: fn())
    monkeypatch.setattr(bridge, "write_yield", lambda *a, **k: None)
    monkeypatch.setattr(bridge, "_ensure", lambda: None)
    with pytest.raises(ExcelBridgeError, match="not numeric"):
        bridge.write_yield_read_pnl("D41", "F44", 3.7)


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


def _parse_watch(*args, **kwargs):
    kwargs.setdefault("watch_cell", "D2")
    return parse_watch_row(*args, **kwargs)


def test_bind_slot_cells_row_41() -> None:
    input_cell, qty_cell, pnl_cell = bind_slot_cells(41, "D", "E", "F", 3)
    assert input_cell == "D41"
    assert qty_cell == "E41"
    assert pnl_cell == "F44"


def test_parse_watch_row_formula_a41() -> None:
    assert (
        _parse_watch("=A41", "국고 25-10", _SLOT_ROWS, _INSTRUMENTS, "트레이딩")
        == 41
    )
    assert _parse_watch("=$A$19", None, _SLOT_ROWS, _INSTRUMENTS) == 19
    assert (
        _parse_watch(
            "=트레이딩!A25", "국고 25-11", _SLOT_ROWS, _INSTRUMENTS, "트레이딩"
        )
        == 25
    )
    assert (
        _parse_watch(
            "='트레이딩'!A46", None, _SLOT_ROWS, _INSTRUMENTS, "트레이딩"
        )
        == 46
    )


def test_parse_watch_row_plain_instrument() -> None:
    assert (
        _parse_watch("국고 25-10", "국고 25-10", _SLOT_ROWS, _INSTRUMENTS)
        == 41
    )


def test_parse_watch_row_rejects_bad_formula() -> None:
    with pytest.raises(ExcelBridgeError):
        _parse_watch("=B41", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        _parse_watch("=A41+0", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        _parse_watch("=A99", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        _parse_watch("=Other!A41", None, _SLOT_ROWS, _INSTRUMENTS, "트레이딩")
    with pytest.raises(ExcelBridgeError):
        _parse_watch("", None, _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        _parse_watch("국고 99-99", "국고 99-99", _SLOT_ROWS, _INSTRUMENTS)
    with pytest.raises(ExcelBridgeError):
        _parse_watch(
            "25-10",
            "25-10",
            _SLOT_ROWS,
            {19: "국고 25-10", 41: "국고 25-10"},
        )


def test_workbook_bind_paths_normalizes() -> None:
    paths = workbook_bind_paths(r"C:/mycode/KBondWatcher/data/sample.xlsm")
    assert paths[0] == r"C:\mycode\KBondWatcher\data\sample.xlsm"
    assert all("\\" in p or p.startswith("C:") for p in paths)


def test_pick_matching_workbook_by_name() -> None:
    wanted = object()
    other = object()
    got = pick_matching_workbook(
        r"C:\mycode\KBondWatcher\data\sample.xlsm",
        [
            ("other.xlsm", r"C:\tmp\other.xlsm", other),
            ("sample.xlsm", r"D:\elsewhere\sample.xlsm", wanted),
        ],
    )
    assert got is wanted


def test_pick_matching_workbook_missing_raises_disconnected() -> None:
    with pytest.raises(ExcelDisconnected, match="is not open"):
        pick_matching_workbook(
            r"C:\mycode\KBondWatcher\data\sample.xlsm",
            [("other.xlsm", r"C:\tmp\other.xlsm", object())],
        )


def test_rot_names_matching_workbook() -> None:
    cfg = r"C:\mycode\KBondWatcher\data\sample.xlsm"
    names = rot_names_matching_workbook(
        cfg,
        [
            "!{00020812-0000-0000-C000-000000000046}",
            r"C:\mycode\KBondWatcher\data\sample.xlsm",
            r"C:\tmp\other.xlsm",
            "",
        ],
    )
    assert names == [r"C:\mycode\KBondWatcher\data\sample.xlsm"]


def test_is_excel_gone_rpc() -> None:
    exc = _FakeComError(-2147023174, "The RPC server is unavailable.")
    assert is_excel_gone(exc) is True
    assert is_excel_busy(exc) is False


def test_is_excel_gone_not_sheet_or_busy() -> None:
    assert is_excel_gone(ExcelBridgeError("Worksheet '트레이딩' not found")) is False
    busy = _FakeComError(-2147417846, "The message filter indicated that the application is busy.")
    assert is_excel_gone(busy) is False
    assert is_excel_gone(ExcelDisconnected("Workbook 'x' is not open in Excel")) is True


def test_call_excel_gone_raises_disconnected() -> None:
    bridge = _dummy_bridge()

    def _gone() -> None:
        raise _FakeComError(-2147023174, "The RPC server is unavailable.")

    with pytest.raises(ExcelDisconnected, match="is not open"):
        bridge._call_excel(_gone)


def test_release_workbook_keeps_com_apartment() -> None:
    bridge = _dummy_bridge()
    sentinel = object()
    bridge._pythoncom = sentinel
    bridge._connected = True
    bridge._app = object()
    bridge._wb = object()
    bridge._ws = object()
    bridge.release_workbook()
    assert bridge._connected is False
    assert bridge._app is None
    assert bridge._wb is None
    assert bridge._ws is None
    assert bridge._pythoncom is sentinel
