from __future__ import annotations

import logging
import math
import re
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional, Sequence, TypeVar, Union

from core.models import AppStatus
from core.trigger import looking_for_from_qty, qty_magnitude

logger = logging.getLogger("kbond_watcher")

XL_DONE = 0
CALC_WAIT_TIMEOUT_SECONDS = 30.0
CALC_POLL_INTERVAL_SECONDS = 0.05
COM_BUSY_RETRY_COUNT = 50
COM_BUSY_RETRY_DELAY_SECONDS = 0.1

_RPC_E_SERVERCALL_RETRYLATER = -2147417846
_RPC_E_CALL_REJECTED = -2147418111
_BUSY_HRESULTS = frozenset(
    {
        _RPC_E_SERVERCALL_RETRYLATER,
        _RPC_E_CALL_REJECTED,
        _RPC_E_SERVERCALL_RETRYLATER & 0xFFFFFFFF,
        _RPC_E_CALL_REJECTED & 0xFFFFFFFF,
    }
)

_GUKGO_PREFIX = re.compile(r"^국고\s*")
_WATCH_A_REF = re.compile(
    r"^=\s*(?:(?:'(?P<qsheet>[^']+)'|(?P<sheet>[^'!]+))!)?\$?(?P<col>[A-Za-z]+)\$?(?P<row>\d+)\s*$",
    re.IGNORECASE,
)
_T = TypeVar("_T")

WATCH_INSTRUMENT_CELL = "D2"
PNL_THRESHOLD_CELL = "E2"


class ExcelBridgeError(RuntimeError):
    pass


class StopRequested(Exception):
    pass


def is_excel_busy(exc: BaseException) -> bool:
    args = getattr(exc, "args", None)
    if not args:
        return False
    hresult = args[0]
    if not isinstance(hresult, int):
        return False
    return hresult in _BUSY_HRESULTS or (hresult - 0x100000000) in _BUSY_HRESULTS


@dataclass(frozen=True)
class InstrumentSlot:
    instrument: str
    row: int
    looking_for: str
    required_side: str
    qty_abs: int
    yield_prefix: float
    input_cell: str
    qty_cell: str
    pnl_cell: str


def to_float(value: Any) -> float:
    if value is None:
        raise ExcelBridgeError("Excel cell value is empty/None")
    if isinstance(value, bool):
        raise ExcelBridgeError(f"unexpected boolean cell value: {value!r}")
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace(",", "")
    if not text:
        raise ExcelBridgeError("Excel cell value is blank")
    try:
        return float(text)
    except ValueError as exc:
        raise ExcelBridgeError(f"cannot convert Excel value to float: {value!r}") from exc


def format_status(status: Union[AppStatus, str]) -> str:
    if isinstance(status, AppStatus):
        return status.value
    return str(status)


def normalize_instrument(raw: Any) -> str:
    text = str(raw or "").strip()
    if not text:
        return ""
    return _GUKGO_PREFIX.sub("", text).strip()


def prefix_from_prev_yield(value: float) -> float:
    return float(math.floor(abs(float(value))))


def bind_slot_cells(
    row: int,
    input_col: str,
    qty_col: str,
    pnl_col: str,
    pnl_row_offset: int,
) -> tuple[str, str, str]:
    return (
        f"{input_col}{row}",
        f"{qty_col}{row}",
        f"{pnl_col}{row + pnl_row_offset}",
    )


def parse_watch_row(
    formula: Any,
    value: Any,
    slot_rows: Sequence[int],
    instruments_by_row: dict[int, str],
    sheet_name: str = "",
    instrument_col: str = "A",
) -> int:
    allowed = {int(r) for r in slot_rows}
    if not allowed:
        raise ExcelBridgeError("slot row allowlist is empty")
    formula_text = str(formula or "").strip()
    col_want = (instrument_col or "A").strip().upper()
    if formula_text.startswith("="):
        matched = _WATCH_A_REF.fullmatch(formula_text)
        if matched is None:
            raise ExcelBridgeError(
                f"{WATCH_INSTRUMENT_CELL} formula is not a single "
                f"{col_want}{{row}} ref: {formula_text!r}"
            )
        col = matched.group("col").upper()
        if col != col_want:
            raise ExcelBridgeError(
                f"{WATCH_INSTRUMENT_CELL} formula must reference column "
                f"{col_want}, got {formula_text!r}"
            )
        sheet = (
            (matched.group("qsheet") or matched.group("sheet") or "")
            .replace("''", "'")
            .strip()
        )
        expected_sheet = (sheet_name or "").strip()
        if sheet and (
            not expected_sheet or sheet.casefold() != expected_sheet.casefold()
        ):
            raise ExcelBridgeError(
                f"{WATCH_INSTRUMENT_CELL} formula must reference the current "
                f"sheet, got {formula_text!r}"
            )
        row = int(matched.group("row"))
        if row not in allowed:
            raise ExcelBridgeError(
                f"{WATCH_INSTRUMENT_CELL} row {row} is not in EXCEL_SLOT_ROWS"
            )
        return row
    target = normalize_instrument(value)
    if not target:
        target = normalize_instrument(formula_text)
    if not target:
        raise ExcelBridgeError(f"{WATCH_INSTRUMENT_CELL} is empty")
    matches = [
        int(row)
        for row in slot_rows
        if int(row) in allowed
        and normalize_instrument(instruments_by_row.get(int(row), "")) == target
    ]
    if len(matches) != 1:
        raise ExcelBridgeError(
            f"{WATCH_INSTRUMENT_CELL} instrument {target!r} matches "
            f"{len(matches)} slot rows"
        )
    return matches[0]


def workbook_matches_open(config_workbook: str, wb_name: str, wb_full_name: str) -> bool:
    configured = (config_workbook or "").strip()
    if not configured:
        return False
    name = (wb_name or "").strip()
    full = (wb_full_name or "").strip()
    cfg_path = Path(configured)
    cfg_lower = configured.replace("/", "\\").lower()
    name_lower = name.lower()
    full_norm = full.replace("/", "\\").lower()
    try:
        cfg_resolved = str(cfg_path.expanduser().resolve()).replace("/", "\\").lower()
    except OSError:
        cfg_resolved = str(cfg_path.expanduser()).replace("/", "\\").lower()
    file_name = cfg_path.name.lower()

    if full_norm and (full_norm == cfg_lower or full_norm == cfg_resolved):
        return True
    if file_name and name_lower == file_name:
        return True
    if name_lower == cfg_lower or (cfg_lower and name_lower.endswith(cfg_lower)):
        return True
    return False


class ExcelBridge:
    def __init__(
        self,
        workbook_name: str,
        sheet_name: str,
        status_cell: str,
        looking_for_cell: str,
        last_quote_cell: str,
        last_pnl_cell: str,
        last_action_cell: str,
        slot_rows: Sequence[int],
        rows_10y: Sequence[int],
        rows_3y: Sequence[int],
        prefix_3y_cell: str,
        prefix_10y_cell: str,
        instrument_col: str,
        qty_col: str,
        input_col: str,
        pnl_col: str,
        pnl_row_offset: int,
        calc_wait_timeout_seconds: float = CALC_WAIT_TIMEOUT_SECONDS,
    ) -> None:
        self.workbook_name = (workbook_name or "").strip()
        self.sheet_name = (sheet_name or "").strip()
        self.status_cell = status_cell
        self.looking_for_cell = looking_for_cell
        self.last_quote_cell = last_quote_cell
        self.last_pnl_cell = last_pnl_cell
        self.last_action_cell = last_action_cell
        self.slot_rows = [int(r) for r in slot_rows]
        self.rows_10y = {int(r) for r in rows_10y}
        self.rows_3y = {int(r) for r in rows_3y}
        self.prefix_3y_cell = prefix_3y_cell
        self.prefix_10y_cell = prefix_10y_cell
        self.instrument_col = instrument_col.strip().upper()
        self.qty_col = qty_col.strip().upper()
        self.input_col = input_col.strip().upper()
        self.pnl_col = pnl_col.strip().upper()
        self.pnl_row_offset = int(pnl_row_offset)
        self.calc_wait_timeout_seconds = float(calc_wait_timeout_seconds)
        self._pythoncom: Any = None
        self._app: Any = None
        self._wb: Any = None
        self._ws: Any = None
        self._connected = False
        self._should_stop: Optional[Callable[[], bool]] = None

    def set_stop_check(self, fn: Optional[Callable[[], bool]]) -> None:
        self._should_stop = fn

    def _call_excel(self, fn: Callable[[], _T], *, check_stop: bool = True) -> _T:
        last: Optional[BaseException] = None
        for _attempt in range(COM_BUSY_RETRY_COUNT):
            if check_stop and self._should_stop is not None and self._should_stop():
                raise StopRequested("stop flag set")
            try:
                return fn()
            except StopRequested:
                raise
            except Exception as exc:
                if is_excel_busy(exc):
                    last = exc
                    time.sleep(COM_BUSY_RETRY_DELAY_SECONDS)
                    continue
                raise
        raise ExcelBridgeError(
            f"Excel busy after {COM_BUSY_RETRY_COUNT} retries: {last}"
        ) from last

    def connect(self) -> None:
        if self._connected:
            return
        try:
            import pythoncom
            import win32com.client
        except ImportError as exc:
            raise ExcelBridgeError("pywin32 is required for Excel COM") from exc

        self._pythoncom = pythoncom
        pythoncom.CoInitialize()
        try:
            self._app = self._call_excel(
                lambda: win32com.client.GetActiveObject("Excel.Application"),
                check_stop=False,
            )
        except ExcelBridgeError:
            raise
        except Exception as exc:
            raise ExcelBridgeError(
                "Failed to connect to running Excel.Application"
            ) from exc

        self._wb = self._resolve_workbook()
        self._ws = self._resolve_worksheet(self._wb)
        self._connected = True
        logger.info(
            "EXCEL_CONNECTED | workbook=%s sheet=%s",
            getattr(self._wb, "Name", "?"),
            getattr(self._ws, "Name", "?"),
        )

    def _resolve_workbook(self) -> Any:
        assert self._app is not None
        if self.workbook_name:
            for i in range(1, self._app.Workbooks.Count + 1):
                wb = self._app.Workbooks(i)
                name = str(wb.Name)
                try:
                    full = str(wb.FullName)
                except Exception:
                    full = ""
                if workbook_matches_open(self.workbook_name, name, full):
                    return wb
            raise ExcelBridgeError(
                f"Workbook '{self.workbook_name}' is not open in Excel"
            )
        wb = self._app.ActiveWorkbook
        if wb is None:
            raise ExcelBridgeError("No ActiveWorkbook available in Excel")
        return wb

    def _resolve_worksheet(self, wb: Any) -> Any:
        if self.sheet_name:
            try:
                return wb.Worksheets(self.sheet_name)
            except Exception as exc:
                raise ExcelBridgeError(
                    f"Worksheet '{self.sheet_name}' not found"
                ) from exc
        ws = wb.ActiveSheet
        if ws is None:
            raise ExcelBridgeError("No ActiveSheet available")
        return ws

    def _ensure(self) -> None:
        if not self._connected or self._ws is None or self._app is None:
            self.connect()

    def read_cell_float(self, addr: str) -> float:
        self._ensure()
        raw = self._call_excel(lambda: self._ws.Range(addr).Value)
        return to_float(raw)

    def read_cell_text(self, addr: str) -> str:
        self._ensure()
        raw = self._call_excel(lambda: self._ws.Range(addr).Value)
        if raw is None:
            return ""
        return str(raw).strip()

    def _yield_prefix_for_row(self, row: int, prefix_3y: float, prefix_10y: float) -> float:
        if row in self.rows_10y:
            return prefix_10y
        if row in self.rows_3y:
            return prefix_3y
        raise ExcelBridgeError(f"row {row} is not mapped to 3Y or 10Y prefix band")

    def load_slots(self) -> tuple[list[InstrumentSlot], str, float]:
        self._ensure()

        def _read() -> tuple[Any, Any, Any, dict[int, Any], dict[int, Any], Any, Any]:
            d2 = self._ws.Range(WATCH_INSTRUMENT_CELL)
            instruments: dict[int, Any] = {}
            qtys: dict[int, Any] = {}
            for row in self.slot_rows:
                instruments[row] = self._ws.Range(
                    f"{self.instrument_col}{row}"
                ).Value
                qtys[row] = self._ws.Range(f"{self.qty_col}{row}").Value
            return (
                d2.Formula,
                d2.Value,
                self._ws.Range(PNL_THRESHOLD_CELL).Value,
                instruments,
                qtys,
                self._ws.Range(self.prefix_3y_cell).Value,
                self._ws.Range(self.prefix_10y_cell).Value,
            )

        (
            formula,
            value,
            threshold_raw,
            instruments,
            qtys,
            prefix_3y_raw,
            prefix_10y_raw,
        ) = self._call_excel(_read)
        threshold = to_float(threshold_raw)
        try:
            prefix_3y = prefix_from_prev_yield(to_float(prefix_3y_raw))
            prefix_10y = prefix_from_prev_yield(to_float(prefix_10y_raw))
        except ExcelBridgeError as exc:
            raise ExcelBridgeError(f"failed to read yield prefixes: {exc}") from exc

        instruments_by_row = {
            int(row): str(raw or "") for row, raw in instruments.items()
        }
        row = parse_watch_row(
            formula,
            value,
            self.slot_rows,
            instruments_by_row,
            sheet_name=self.sheet_name,
            instrument_col=self.instrument_col,
        )
        instrument = normalize_instrument(instruments_by_row.get(row, ""))
        if not instrument:
            raise ExcelBridgeError(f"{self.instrument_col}{row} is empty")
        try:
            qty = to_float(qtys[row])
            looking_for, required_side = looking_for_from_qty(qty)
            qty_abs = qty_magnitude(qty)
        except (ExcelBridgeError, ValueError) as exc:
            raise ExcelBridgeError(
                f"{self.qty_col}{row} must be a non-zero integer"
            ) from exc
        yield_prefix = self._yield_prefix_for_row(row, prefix_3y, prefix_10y)
        input_cell, qty_cell, pnl_cell = bind_slot_cells(
            row,
            self.input_col,
            self.qty_col,
            self.pnl_col,
            self.pnl_row_offset,
        )
        slot = InstrumentSlot(
            instrument=instrument,
            row=row,
            looking_for=looking_for,
            required_side=required_side,
            qty_abs=qty_abs,
            yield_prefix=yield_prefix,
            input_cell=input_cell,
            qty_cell=qty_cell,
            pnl_cell=pnl_cell,
        )
        logger.info(
            "SLOTS_LOADED | count=1 looking_for=%s row=%s instrument=%s "
            "qty=%s threshold=%s prefixes=3Y:%s 10Y:%s",
            looking_for,
            row,
            instrument,
            qty_abs,
            threshold,
            prefix_3y,
            prefix_10y,
        )
        return [slot], looking_for, threshold

    def write_yield(self, input_cell: str, yield_value: float) -> None:
        self._ensure()
        value = float(yield_value)

        def _write() -> None:
            self._ws.Range(input_cell).Value = value

        self._call_excel(_write)
        logger.info("EXCEL_WRITE | %s=%s", input_cell, yield_value)

    def read_pnl(self, pnl_cell: str) -> float:
        self._ensure()
        raw = self._call_excel(lambda: self._ws.Range(pnl_cell).Value)
        pnl = to_float(raw)
        logger.info("PNL | %s=%s", pnl_cell, pnl)
        return pnl

    def _wait_calculation_done(self) -> None:
        assert self._app is not None
        deadline = time.monotonic() + self.calc_wait_timeout_seconds
        state = None
        while True:
            try:
                state = int(self._call_excel(lambda: self._app.CalculationState))
            except StopRequested:
                raise
            except ExcelBridgeError:
                raise
            except Exception as exc:
                raise ExcelBridgeError(
                    f"Failed to read Excel CalculationState: {exc}"
                ) from exc
            if state == XL_DONE:
                return
            if time.monotonic() >= deadline:
                raise ExcelBridgeError(
                    f"Excel calculation did not finish within "
                    f"{self.calc_wait_timeout_seconds:.1f}s "
                    f"(CalculationState={state})"
                )
            time.sleep(CALC_POLL_INTERVAL_SECONDS)

    def write_yield_read_pnl(
        self,
        input_cell: str,
        pnl_cell: str,
        yield_value: float,
    ) -> float:
        self.write_yield(input_cell, yield_value)
        self._wait_calculation_done()
        return self.read_pnl(pnl_cell)

    def update_status(
        self,
        status: Union[AppStatus, str],
        looking_for: Optional[str] = None,
        last_quote: Optional[str] = None,
        last_pnl: Optional[float] = None,
        last_action: Optional[str] = None,
    ) -> None:
        try:
            self._ensure()

            def _write() -> None:
                self._ws.Range(self.status_cell).Value = format_status(status)
                if looking_for is not None:
                    self._ws.Range(self.looking_for_cell).Value = looking_for
                if last_quote is not None:
                    self._ws.Range(self.last_quote_cell).Value = last_quote
                if last_pnl is not None:
                    self._ws.Range(self.last_pnl_cell).Value = float(last_pnl)
                if last_action is not None:
                    self._ws.Range(self.last_action_cell).Value = last_action

            self._call_excel(_write, check_stop=False)
        except StopRequested:
            raise
        except Exception as exc:
            logger.warning("Excel status update failed: %s", exc)

    def close(self) -> None:
        if self._pythoncom is not None:
            try:
                self._pythoncom.CoUninitialize()
            except Exception:
                pass
        self._connected = False
        self._app = None
        self._wb = None
        self._ws = None
