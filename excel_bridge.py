from __future__ import annotations

import logging
from typing import Any, Optional, Union

from models import AppStatus

logger = logging.getLogger("kbond_watcher")


class ExcelBridgeError(RuntimeError):
    pass


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


class ExcelBridge:
    def __init__(
        self,
        workbook_name: str,
        sheet_name: str,
        input_cell: str,
        pnl_cell: str,
        status_cell: str,
        last_quote_cell: str,
        last_pnl_cell: str,
        last_action_cell: str,
    ) -> None:
        self.workbook_name = (workbook_name or "").strip()
        self.sheet_name = (sheet_name or "").strip()
        self.input_cell = input_cell
        self.pnl_cell = pnl_cell
        self.status_cell = status_cell
        self.last_quote_cell = last_quote_cell
        self.last_pnl_cell = last_pnl_cell
        self.last_action_cell = last_action_cell
        self._pythoncom: Any = None
        self._app: Any = None
        self._wb: Any = None
        self._ws: Any = None
        self._connected = False

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
            self._app = win32com.client.GetActiveObject("Excel.Application")
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
            needle = self.workbook_name.lower()
            for i in range(1, self._app.Workbooks.Count + 1):
                wb = self._app.Workbooks(i)
                name = str(wb.Name)
                if name.lower() == needle or name.lower().endswith(needle):
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

    def write_yield(self, yield_value: float) -> None:
        self._ensure()
        self._ws.Range(self.input_cell).Value = float(yield_value)
        logger.info("EXCEL_WRITE | %s=%s", self.input_cell, yield_value)

    def read_pnl(self) -> float:
        self._ensure()
        raw = self._ws.Range(self.pnl_cell).Value
        pnl = to_float(raw)
        logger.info("PNL | %s=%s", self.pnl_cell, pnl)
        return pnl

    def write_yield_read_pnl(self, yield_value: float) -> float:
        self.write_yield(yield_value)
        return self.read_pnl()

    def update_status(
        self,
        status: Union[AppStatus, str],
        last_quote: Optional[str] = None,
        last_pnl: Optional[float] = None,
        last_action: Optional[str] = None,
    ) -> None:
        try:
            self._ensure()
            self._ws.Range(self.status_cell).Value = format_status(status)
            if last_quote is not None:
                self._ws.Range(self.last_quote_cell).Value = last_quote
            if last_pnl is not None:
                self._ws.Range(self.last_pnl_cell).Value = float(last_pnl)
            if last_action is not None:
                self._ws.Range(self.last_action_cell).Value = last_action
        except Exception as exc:
            logger.warning("Excel status update failed: %s", exc)

    def _ensure(self) -> None:
        if not self._connected or self._ws is None or self._app is None:
            self.connect()

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
