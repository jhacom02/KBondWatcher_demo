"""Load and validate config.env for the FORESTBOND watcher."""

from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from dotenv import load_dotenv


class ConfigError(ValueError):
    """Raised when config.env is missing or invalid."""


def _parse_bool(value: str, key: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off", ""}:
        return False
    raise ConfigError(f"{key} must be a boolean, got {value!r}")


def _parse_float(value: str, key: str) -> float:
    try:
        return float(value.replace(",", ""))
    except (TypeError, ValueError) as exc:
        raise ConfigError(f"{key} must be a number, got {value!r}") from exc


def _parse_int(value: str, key: str) -> int:
    try:
        return int(float(value.replace(",", "")))
    except (TypeError, ValueError) as exc:
        raise ConfigError(f"{key} must be an integer, got {value!r}") from exc


@dataclass(frozen=True)
class Config:
    # FORESTBOND
    target: str
    chrome_title: str
    yield_prefix: float
    required_side: str
    poll_interval_ms: int
    process_existing_on_start: bool

    # Excel
    excel_workbook: str
    excel_sheet: str
    excel_input_cell: str
    excel_pnl_cell: str
    excel_status_cell: str
    excel_last_quote_cell: str
    excel_last_pnl_cell: str
    excel_last_action_cell: str
    pnl_threshold: float

    # K-Bond
    kbond_process_name: str
    kbond_pid: int
    kbond_window_title_contains: str
    win_x: float
    win_y: float
    send_text: str
    send_enter: bool

    stop_flag_path: Path
    log_level: str
    log_path: Path

    config_path: Path

    @classmethod
    def load(cls, config_path: str | Path) -> "Config":
        path = Path(config_path).expanduser().resolve()
        if not path.is_file():
            raise ConfigError(f"config file not found: {path}")

        load_dotenv(path, override=True)
        # Also read file directly so comments/order do not matter and missing
        # keys fall back to defaults even if OS env already has unrelated vars.
        file_values: dict[str, str] = {}
        for line in path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, _, val = stripped.partition("=")
            file_values[key.strip()] = val.strip().strip('"').strip("'")

        def get(key: str, default: str = "") -> str:
            if key in file_values:
                return file_values[key]
            return os.getenv(key, default)

        target = get("TARGET", "25-11")
        if not target:
            raise ConfigError("TARGET must not be empty")

        chrome_title = get("CHROME_TITLE", "FORESTBOND")
        if not chrome_title:
            raise ConfigError("CHROME_TITLE must not be empty")

        yield_prefix = _parse_float(get("YIELD_PREFIX", "4"), "YIELD_PREFIX")

        required_side = get("REQUIRED_SIDE", "ANY").upper()
        if required_side not in {"ANY", "BUY", "SELL"}:
            raise ConfigError("REQUIRED_SIDE must be ANY, BUY, or SELL")

        poll_interval_ms = _parse_int(get("POLL_INTERVAL_MS", "300"), "POLL_INTERVAL_MS")
        if poll_interval_ms < 100:
            raise ConfigError("POLL_INTERVAL_MS must be >= 100")

        process_existing_on_start = _parse_bool(
            get("PROCESS_EXISTING_ON_START", "true"),
            "PROCESS_EXISTING_ON_START",
        )

        excel_workbook = get("EXCEL_WORKBOOK", "sample.xlsx")
        excel_sheet = get("EXCEL_SHEET", "")
        excel_input_cell = get("EXCEL_INPUT_CELL", "D19")
        excel_pnl_cell = get("EXCEL_PNL_CELL", "F22")
        excel_status_cell = get("EXCEL_STATUS_CELL", "J16")
        excel_last_quote_cell = get("EXCEL_LAST_QUOTE_CELL", "J17")
        excel_last_pnl_cell = get("EXCEL_LAST_PNL_CELL", "J18")
        excel_last_action_cell = get("EXCEL_LAST_ACTION_CELL", "J19")

        for cell_key, cell_val in [
            ("EXCEL_INPUT_CELL", excel_input_cell),
            ("EXCEL_PNL_CELL", excel_pnl_cell),
            ("EXCEL_STATUS_CELL", excel_status_cell),
            ("EXCEL_LAST_QUOTE_CELL", excel_last_quote_cell),
            ("EXCEL_LAST_PNL_CELL", excel_last_pnl_cell),
            ("EXCEL_LAST_ACTION_CELL", excel_last_action_cell),
        ]:
            if not cell_val:
                raise ConfigError(f"{cell_key} must not be empty")

        pnl_threshold = _parse_float(get("PNL_THRESHOLD", "1000000"), "PNL_THRESHOLD")

        kbond_process_name = get("KBOND_PROCESS_NAME", "axis.exe")
        if not kbond_process_name:
            raise ConfigError("KBOND_PROCESS_NAME must not be empty")
        if not kbond_process_name.lower().endswith(".exe"):
            kbond_process_name = f"{kbond_process_name}.exe"

        kbond_pid = _parse_int(get("KBOND_PID", "0"), "KBOND_PID")
        if kbond_pid < 0:
            raise ConfigError("KBOND_PID must be >= 0")

        kbond_window_title_contains = get("KBOND_WINDOW_TITLE_CONTAINS", "")

        win_x = _parse_float(get("WIN_X", "0.5"), "WIN_X")
        win_y = _parse_float(get("WIN_Y", "0.4"), "WIN_Y")
        if not (0.0 <= win_x <= 1.0):
            raise ConfigError("WIN_X must be between 0 and 1 inclusive")
        if not (0.0 <= win_y <= 1.0):
            raise ConfigError("WIN_Y must be between 0 and 1 inclusive")

        send_text = get("SEND_TEXT", "ㅎㅈ")
        if not send_text:
            raise ConfigError("SEND_TEXT must not be empty")

        send_enter = _parse_bool(get("SEND_ENTER", "false"), "SEND_ENTER")

        stop_flag_raw = get("STOP_FLAG_PATH", r"C:\temp\forestbond_watcher.stop")
        stop_flag_path = Path(stop_flag_raw).expanduser()

        log_level = get("LOG_LEVEL", "INFO").upper()
        log_path_raw = get("LOG_PATH", r"logs\watcher.log")
        log_path = Path(log_path_raw)
        if not log_path.is_absolute():
            log_path = path.parent / log_path

        return cls(
            target=target,
            chrome_title=chrome_title,
            yield_prefix=yield_prefix,
            required_side=required_side,
            poll_interval_ms=poll_interval_ms,
            process_existing_on_start=process_existing_on_start,
            excel_workbook=excel_workbook,
            excel_sheet=excel_sheet,
            excel_input_cell=excel_input_cell,
            excel_pnl_cell=excel_pnl_cell,
            excel_status_cell=excel_status_cell,
            excel_last_quote_cell=excel_last_quote_cell,
            excel_last_pnl_cell=excel_last_pnl_cell,
            excel_last_action_cell=excel_last_action_cell,
            pnl_threshold=pnl_threshold,
            kbond_process_name=kbond_process_name,
            kbond_pid=kbond_pid,
            kbond_window_title_contains=kbond_window_title_contains,
            win_x=win_x,
            win_y=win_y,
            send_text=send_text,
            send_enter=send_enter,
            stop_flag_path=stop_flag_path,
            log_level=log_level,
            log_path=log_path,
            config_path=path,
        )
