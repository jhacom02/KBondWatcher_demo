from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path

from dotenv import load_dotenv


class ConfigError(ValueError):
    pass


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


def _require_ratio(value: float, key: str) -> float:
    if not (0.0 <= value <= 1.0):
        raise ConfigError(f"{key} must be between 0 and 1 inclusive")
    return value


@dataclass(frozen=True)
class Config:
    target: str
    source_window_title: str
    source_process_name: str
    yield_prefix: float
    required_side: str
    poll_interval_ms: int
    process_existing_on_start: bool

    excel_workbook: str
    excel_sheet: str
    excel_input_cell: str
    excel_pnl_cell: str
    excel_status_cell: str
    excel_last_quote_cell: str
    excel_last_pnl_cell: str
    excel_last_action_cell: str
    pnl_threshold: float

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

        target = require("TARGET")
        if not target:
            raise ConfigError("TARGET must not be empty")

        source_window_title = require("SOURCE_WINDOW_TITLE")
        if not source_window_title:
            raise ConfigError("SOURCE_WINDOW_TITLE must not be empty")
        source_process_name = optional("SOURCE_PROCESS_NAME").strip()

        yield_prefix = _parse_float(require("YIELD_PREFIX"), "YIELD_PREFIX")

        required_side = require("REQUIRED_SIDE").upper()
        if required_side not in {"ANY", "BUY", "SELL"}:
            raise ConfigError("REQUIRED_SIDE must be ANY, BUY, or SELL")

        poll_interval_ms = _parse_int(require("POLL_INTERVAL_MS"), "POLL_INTERVAL_MS")
        if poll_interval_ms < 100:
            raise ConfigError("POLL_INTERVAL_MS must be >= 100")

        process_existing_on_start = _parse_bool(
            require("PROCESS_EXISTING_ON_START"),
            "PROCESS_EXISTING_ON_START",
        )

        excel_workbook = require("EXCEL_WORKBOOK")
        if not excel_workbook:
            raise ConfigError("EXCEL_WORKBOOK must not be empty")
        excel_sheet = optional("EXCEL_SHEET")
        excel_input_cell = require("EXCEL_INPUT_CELL")
        excel_pnl_cell = require("EXCEL_PNL_CELL")
        excel_status_cell = require("EXCEL_STATUS_CELL")
        excel_last_quote_cell = require("EXCEL_LAST_QUOTE_CELL")
        excel_last_pnl_cell = require("EXCEL_LAST_PNL_CELL")
        excel_last_action_cell = require("EXCEL_LAST_ACTION_CELL")
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

        pnl_threshold = _parse_float(require("PNL_THRESHOLD"), "PNL_THRESHOLD")

        send_process_name = require("SEND_PROCESS_NAME")
        if not send_process_name:
            raise ConfigError("SEND_PROCESS_NAME must not be empty")
        send_window_title = require("SEND_WINDOW_TITLE")
        if not send_window_title:
            raise ConfigError("SEND_WINDOW_TITLE must not be empty")
        message_template = require("MESSAGE_TEMPLATE")
        if not message_template:
            raise ConfigError("MESSAGE_TEMPLATE must not be empty")

        send_input_x = _require_ratio(
            _parse_float(require("SEND_INPUT_X"), "SEND_INPUT_X"),
            "SEND_INPUT_X",
        )
        send_input_y = _require_ratio(
            _parse_float(require("SEND_INPUT_Y"), "SEND_INPUT_Y"),
            "SEND_INPUT_Y",
        )

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
            target=target,
            source_window_title=source_window_title,
            source_process_name=source_process_name,
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
        )
