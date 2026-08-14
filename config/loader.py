from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


class ConfigError(ValueError):
    pass


KBOND_PROCESS = "KBondMessenger.exe"
KBOND_TITLE = "K-Bond"
FORESTBOND_TITLE = "FORESTBOND"
NOTEPAD_PROCESS = "notepad.exe"
NOTEPAD_TITLE = "메모장"


def _parse_bool(value: str, key: str) -> bool:
    normalized = value.strip().lower()
    if normalized in {"1", "true", "yes", "y", "on"}:
        return True
    if normalized in {"0", "false", "no", "n", "off"}:
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


def _parse_column(value: str, key: str) -> str:
    col = (value or "").strip().upper()
    if not col or not col.isalpha():
        raise ConfigError(f"{key} must be a column letter like A or AA, got {value!r}")
    return col


def _parse_int_list(value: str, key: str) -> list[int]:
    parts = [p.strip() for p in value.replace(";", ",").split(",") if p.strip()]
    if not parts:
        raise ConfigError(f"{key} must not be empty")
    out: list[int] = []
    for part in parts:
        out.append(_parse_int(part, key))
    return out


def _parse_ratio(value: str, key: str) -> float:
    ratio = _parse_float(value, key)
    if not 0.0 <= ratio <= 1.0:
        raise ConfigError(f"{key} must be between 0 and 1, got {value!r}")
    return ratio


def mode_presets(mode: int) -> tuple[str, str, str, str]:
    if mode == 1:
        return (
            KBOND_TITLE,
            KBOND_PROCESS,
            KBOND_PROCESS,
            KBOND_TITLE,
        )
    if mode == 2:
        return (
            KBOND_TITLE,
            KBOND_PROCESS,
            NOTEPAD_PROCESS,
            NOTEPAD_TITLE,
        )
    if mode == 3:
        return (
            FORESTBOND_TITLE,
            "",
            NOTEPAD_PROCESS,
            NOTEPAD_TITLE,
        )
    raise ConfigError(f"MODE must be 1, 2, or 3, got {mode}")


@dataclass(frozen=True)
class Config:
    mode: int
    source_window_title: str
    source_process_name: str
    poll_interval_ms: int
    process_existing_on_start: bool

    excel_workbook: str
    excel_sheet: str
    excel_slot_rows: tuple[int, ...]
    excel_rows_10y: tuple[int, ...]
    excel_rows_3y: tuple[int, ...]
    excel_instrument_col: str
    excel_qty_col: str
    excel_input_col: str
    excel_pnl_col: str
    excel_pnl_row_offset: int
    excel_prefix_3y_cell: str
    excel_prefix_10y_cell: str
    excel_watch_cell: str
    excel_pnl_threshold_cell: str
    excel_pnl_sanity_band: float
    excel_status_cell: str
    excel_looking_for_cell: str
    excel_last_quote_cell: str
    excel_last_pnl_cell: str
    excel_last_action_cell: str

    send_process_name: str
    send_window_title: str
    message_template: str
    send_input_x: float
    send_input_y: float
    send_foreground_retry_pause_seconds: float
    send_activate_show_pause_seconds: float
    send_after_activate_pause_seconds: float
    send_input_click_pause_seconds: float
    send_paste_pause_seconds: float
    send_send_pause_seconds: float

    stop_flag_path: Path
    log_level: str
    log_path: Path
    config_path: Path
    sent_after: str

    @classmethod
    def load(cls, config_path: str | Path) -> "Config":
        path = Path(config_path).expanduser().resolve()
        if not path.is_file():
            raise ConfigError(f"config file not found: {path}")

        load_dotenv(path, override=True)
        file_values: dict[str, str] = {}
        for line in path.read_text(encoding="utf-8").splitlines():
            stripped = line.strip()
            if not stripped or stripped.startswith("#") or "=" not in stripped:
                continue
            key, _, val = stripped.partition("=")
            file_values[key.strip()] = val.strip().strip('"').strip("'")

        def require(key: str) -> str:
            if key in file_values:
                return file_values[key]
            env_val = os.getenv(key)
            if env_val is not None:
                return env_val
            raise ConfigError(f"missing required config key: {key}")

        def optional(key: str) -> str:
            if key in file_values:
                return file_values[key]
            return os.getenv(key, "")

        mode = _parse_int(require("MODE"), "MODE")
        (
            source_window_title,
            source_process_name,
            send_process_name,
            send_window_title,
        ) = mode_presets(mode)

        if mode in (1, 2):
            chat_title = require("KBOND_CHAT_TITLE").strip()
            if not chat_title:
                raise ConfigError("KBOND_CHAT_TITLE must not be empty")
            source_window_title = chat_title
            if mode == 1:
                send_window_title = chat_title

        if mode == 1:
            send_input_x = _parse_ratio(require("SEND_INPUT_X_M1"), "SEND_INPUT_X_M1")
            send_input_y = _parse_ratio(require("SEND_INPUT_Y_M1"), "SEND_INPUT_Y_M1")
        else:
            send_input_x = _parse_ratio(require("SEND_INPUT_X_M23"), "SEND_INPUT_X_M23")
            send_input_y = _parse_ratio(require("SEND_INPUT_Y_M23"), "SEND_INPUT_Y_M23")

        poll_interval_ms = _parse_int(require("POLL_INTERVAL_MS"), "POLL_INTERVAL_MS")
        if poll_interval_ms < 100:
            raise ConfigError("POLL_INTERVAL_MS must be >= 100")

        process_existing_on_start = _parse_bool(
            require("PROCESS_EXISTING_ON_START"),
            "PROCESS_EXISTING_ON_START",
        )

        sent_after = require("SENT_AFTER").strip().lower()
        if sent_after not in {"exit", "loop"}:
            raise ConfigError("SENT_AFTER must be exit or loop")

        excel_workbook = require("EXCEL_WORKBOOK")
        if not excel_workbook:
            raise ConfigError("EXCEL_WORKBOOK must not be empty")
        excel_sheet = optional("EXCEL_SHEET")
        excel_slot_rows = tuple(
            _parse_int_list(require("EXCEL_SLOT_ROWS"), "EXCEL_SLOT_ROWS")
        )
        excel_rows_10y = tuple(
            _parse_int_list(require("EXCEL_ROWS_10Y"), "EXCEL_ROWS_10Y")
        )
        excel_rows_3y = tuple(
            _parse_int_list(require("EXCEL_ROWS_3Y"), "EXCEL_ROWS_3Y")
        )
        excel_instrument_col = _parse_column(
            require("EXCEL_INSTRUMENT_COL"), "EXCEL_INSTRUMENT_COL"
        )
        excel_qty_col = _parse_column(require("EXCEL_QTY_COL"), "EXCEL_QTY_COL")
        excel_input_col = _parse_column(require("EXCEL_INPUT_COL"), "EXCEL_INPUT_COL")
        excel_pnl_col = _parse_column(require("EXCEL_PNL_COL"), "EXCEL_PNL_COL")
        excel_pnl_row_offset = _parse_int(
            require("EXCEL_PNL_ROW_OFFSET"), "EXCEL_PNL_ROW_OFFSET"
        )
        if excel_pnl_row_offset < 0:
            raise ConfigError("EXCEL_PNL_ROW_OFFSET must be >= 0")
        excel_pnl_sanity_band = _parse_float(
            require("EXCEL_PNL_SANITY_BAND"), "EXCEL_PNL_SANITY_BAND"
        )
        if excel_pnl_sanity_band <= 0:
            raise ConfigError("EXCEL_PNL_SANITY_BAND must be > 0")
        excel_prefix_3y_cell = require("EXCEL_PREFIX_3Y_CELL")
        excel_prefix_10y_cell = require("EXCEL_PREFIX_10Y_CELL")
        excel_watch_cell = require("EXCEL_WATCH_CELL")
        excel_pnl_threshold_cell = require("EXCEL_PNL_THRESHOLD_CELL")
        excel_status_cell = require("EXCEL_STATUS_CELL")
        excel_looking_for_cell = require("EXCEL_LOOKING_FOR_CELL")
        excel_last_quote_cell = require("EXCEL_LAST_QUOTE_CELL")
        excel_last_pnl_cell = require("EXCEL_LAST_PNL_CELL")
        excel_last_action_cell = require("EXCEL_LAST_ACTION_CELL")
        for cell_key, cell_val in [
            ("EXCEL_PREFIX_3Y_CELL", excel_prefix_3y_cell),
            ("EXCEL_PREFIX_10Y_CELL", excel_prefix_10y_cell),
            ("EXCEL_WATCH_CELL", excel_watch_cell),
            ("EXCEL_PNL_THRESHOLD_CELL", excel_pnl_threshold_cell),
            ("EXCEL_STATUS_CELL", excel_status_cell),
            ("EXCEL_LOOKING_FOR_CELL", excel_looking_for_cell),
            ("EXCEL_LAST_QUOTE_CELL", excel_last_quote_cell),
            ("EXCEL_LAST_PNL_CELL", excel_last_pnl_cell),
            ("EXCEL_LAST_ACTION_CELL", excel_last_action_cell),
        ]:
            if not cell_val:
                raise ConfigError(f"{cell_key} must not be empty")

        mapped = set(excel_rows_10y) | set(excel_rows_3y)
        for row in excel_slot_rows:
            if row not in mapped:
                raise ConfigError(
                    f"EXCEL_SLOT_ROWS contains {row} not in EXCEL_ROWS_10Y/3Y"
                )

        message_template = require("MESSAGE_TEMPLATE")
        if not message_template:
            raise ConfigError("MESSAGE_TEMPLATE must not be empty")

        send_foreground_retry_pause_seconds = _parse_float(
            require("SEND_FOREGROUND_RETRY_PAUSE_SECONDS"),
            "SEND_FOREGROUND_RETRY_PAUSE_SECONDS",
        )
        send_activate_show_pause_seconds = _parse_float(
            require("SEND_ACTIVATE_SHOW_PAUSE_SECONDS"),
            "SEND_ACTIVATE_SHOW_PAUSE_SECONDS",
        )
        send_after_activate_pause_seconds = _parse_float(
            require("SEND_AFTER_ACTIVATE_PAUSE_SECONDS"),
            "SEND_AFTER_ACTIVATE_PAUSE_SECONDS",
        )
        send_input_click_pause_seconds = _parse_float(
            require("SEND_INPUT_CLICK_PAUSE_SECONDS"),
            "SEND_INPUT_CLICK_PAUSE_SECONDS",
        )
        send_paste_pause_seconds = _parse_float(
            require("SEND_PASTE_PAUSE_SECONDS"), "SEND_PASTE_PAUSE_SECONDS"
        )
        send_send_pause_seconds = _parse_float(
            require("SEND_SEND_PAUSE_SECONDS"), "SEND_SEND_PAUSE_SECONDS"
        )

        stop_flag_raw = require("STOP_FLAG_PATH")
        if not stop_flag_raw:
            raise ConfigError("STOP_FLAG_PATH must not be empty")
        stop_flag_path = Path(stop_flag_raw).expanduser()

        log_level = require("LOG_LEVEL").upper()
        log_path_raw = require("LOG_PATH")
        if not log_path_raw:
            raise ConfigError("LOG_PATH must not be empty")
        log_path = Path(log_path_raw)
        if not log_path.is_absolute():
            log_path = path.parent / log_path

        return cls(
            mode=mode,
            source_window_title=source_window_title,
            source_process_name=source_process_name,
            poll_interval_ms=poll_interval_ms,
            process_existing_on_start=process_existing_on_start,
            excel_workbook=excel_workbook,
            excel_sheet=excel_sheet,
            excel_slot_rows=excel_slot_rows,
            excel_rows_10y=excel_rows_10y,
            excel_rows_3y=excel_rows_3y,
            excel_instrument_col=excel_instrument_col,
            excel_qty_col=excel_qty_col,
            excel_input_col=excel_input_col,
            excel_pnl_col=excel_pnl_col,
            excel_pnl_row_offset=excel_pnl_row_offset,
            excel_prefix_3y_cell=excel_prefix_3y_cell,
            excel_prefix_10y_cell=excel_prefix_10y_cell,
            excel_watch_cell=excel_watch_cell,
            excel_pnl_threshold_cell=excel_pnl_threshold_cell,
            excel_pnl_sanity_band=excel_pnl_sanity_band,
            excel_status_cell=excel_status_cell,
            excel_looking_for_cell=excel_looking_for_cell,
            excel_last_quote_cell=excel_last_quote_cell,
            excel_last_pnl_cell=excel_last_pnl_cell,
            excel_last_action_cell=excel_last_action_cell,
            send_process_name=send_process_name,
            send_window_title=send_window_title,
            message_template=message_template,
            send_input_x=send_input_x,
            send_input_y=send_input_y,
            send_foreground_retry_pause_seconds=send_foreground_retry_pause_seconds,
            send_activate_show_pause_seconds=send_activate_show_pause_seconds,
            send_after_activate_pause_seconds=send_after_activate_pause_seconds,
            send_input_click_pause_seconds=send_input_click_pause_seconds,
            send_paste_pause_seconds=send_paste_pause_seconds,
            send_send_pause_seconds=send_send_pause_seconds,
            stop_flag_path=stop_flag_path,
            log_level=log_level,
            log_path=log_path,
            config_path=path,
            sent_after=sent_after,
        )
