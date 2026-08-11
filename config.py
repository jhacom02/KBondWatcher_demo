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
    chrome_title: str
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

    kakao_process_name: str
    kakao_window_class: str
    kakao_main_title: str
    kakao_room_name: str
    message_template: str
    kakao_chat_tab_x: float
    kakao_chat_tab_y: float
    kakao_input_x: float
    kakao_input_y: float
    kakao_room_window_wait_seconds: float
    kakao_foreground_retry_pause_seconds: float
    kakao_window_poll_interval_seconds: float
    kakao_activate_show_pause_seconds: float
    kakao_after_activate_pause_seconds: float
    kakao_chat_tab_pause_seconds: float
    kakao_search_open_pause_seconds: float
    kakao_search_reset_pause_seconds: float
    kakao_search_paste_pause_seconds: float
    kakao_room_enter_pause_seconds: float
    kakao_input_click_pause_seconds: float
    kakao_paste_pause_seconds: float
    kakao_send_pause_seconds: float
    kakao_search_clear_backspace_count: int

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

        chrome_title = require("CHROME_TITLE")
        if not chrome_title:
            raise ConfigError("CHROME_TITLE must not be empty")

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

        kakao_process_name = require("KAKAO_PROCESS_NAME")
        if not kakao_process_name:
            raise ConfigError("KAKAO_PROCESS_NAME must not be empty")
        kakao_window_class = require("KAKAO_WINDOW_CLASS")
        if not kakao_window_class:
            raise ConfigError("KAKAO_WINDOW_CLASS must not be empty")
        kakao_main_title = require("KAKAO_MAIN_TITLE")
        if not kakao_main_title:
            raise ConfigError("KAKAO_MAIN_TITLE must not be empty")
        kakao_room_name = require("KAKAO_ROOM_NAME")
        if not kakao_room_name:
            raise ConfigError("KAKAO_ROOM_NAME must not be empty")
        message_template = require("MESSAGE_TEMPLATE")
        if not message_template:
            raise ConfigError("MESSAGE_TEMPLATE must not be empty")

        kakao_chat_tab_x = _require_ratio(
            _parse_float(require("KAKAO_CHAT_TAB_X"), "KAKAO_CHAT_TAB_X"),
            "KAKAO_CHAT_TAB_X",
        )
        kakao_chat_tab_y = _require_ratio(
            _parse_float(require("KAKAO_CHAT_TAB_Y"), "KAKAO_CHAT_TAB_Y"),
            "KAKAO_CHAT_TAB_Y",
        )
        kakao_input_x = _require_ratio(
            _parse_float(require("KAKAO_INPUT_X"), "KAKAO_INPUT_X"),
            "KAKAO_INPUT_X",
        )
        kakao_input_y = _require_ratio(
            _parse_float(require("KAKAO_INPUT_Y"), "KAKAO_INPUT_Y"),
            "KAKAO_INPUT_Y",
        )

        kakao_room_window_wait_seconds = _parse_float(
            require("KAKAO_ROOM_WINDOW_WAIT_SECONDS"), "KAKAO_ROOM_WINDOW_WAIT_SECONDS"
        )
        kakao_foreground_retry_pause_seconds = _parse_float(
            require("KAKAO_FOREGROUND_RETRY_PAUSE_SECONDS"),
            "KAKAO_FOREGROUND_RETRY_PAUSE_SECONDS",
        )
        kakao_window_poll_interval_seconds = _parse_float(
            require("KAKAO_WINDOW_POLL_INTERVAL_SECONDS"),
            "KAKAO_WINDOW_POLL_INTERVAL_SECONDS",
        )
        kakao_activate_show_pause_seconds = _parse_float(
            require("KAKAO_ACTIVATE_SHOW_PAUSE_SECONDS"),
            "KAKAO_ACTIVATE_SHOW_PAUSE_SECONDS",
        )
        kakao_after_activate_pause_seconds = _parse_float(
            require("KAKAO_AFTER_ACTIVATE_PAUSE_SECONDS"),
            "KAKAO_AFTER_ACTIVATE_PAUSE_SECONDS",
        )
        kakao_chat_tab_pause_seconds = _parse_float(
            require("KAKAO_CHAT_TAB_PAUSE_SECONDS"), "KAKAO_CHAT_TAB_PAUSE_SECONDS"
        )
        kakao_search_open_pause_seconds = _parse_float(
            require("KAKAO_SEARCH_OPEN_PAUSE_SECONDS"),
            "KAKAO_SEARCH_OPEN_PAUSE_SECONDS",
        )
        kakao_search_reset_pause_seconds = _parse_float(
            require("KAKAO_SEARCH_RESET_PAUSE_SECONDS"),
            "KAKAO_SEARCH_RESET_PAUSE_SECONDS",
        )
        kakao_search_paste_pause_seconds = _parse_float(
            require("KAKAO_SEARCH_PASTE_PAUSE_SECONDS"),
            "KAKAO_SEARCH_PASTE_PAUSE_SECONDS",
        )
        kakao_room_enter_pause_seconds = _parse_float(
            require("KAKAO_ROOM_ENTER_PAUSE_SECONDS"),
            "KAKAO_ROOM_ENTER_PAUSE_SECONDS",
        )
        kakao_input_click_pause_seconds = _parse_float(
            require("KAKAO_INPUT_CLICK_PAUSE_SECONDS"),
            "KAKAO_INPUT_CLICK_PAUSE_SECONDS",
        )
        kakao_paste_pause_seconds = _parse_float(
            require("KAKAO_PASTE_PAUSE_SECONDS"), "KAKAO_PASTE_PAUSE_SECONDS"
        )
        kakao_send_pause_seconds = _parse_float(
            require("KAKAO_SEND_PAUSE_SECONDS"), "KAKAO_SEND_PAUSE_SECONDS"
        )
        kakao_search_clear_backspace_count = _parse_int(
            require("KAKAO_SEARCH_CLEAR_BACKSPACE_COUNT"),
            "KAKAO_SEARCH_CLEAR_BACKSPACE_COUNT",
        )
        if kakao_search_clear_backspace_count < 0:
            raise ConfigError("KAKAO_SEARCH_CLEAR_BACKSPACE_COUNT must be >= 0")

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
            kakao_process_name=kakao_process_name,
            kakao_window_class=kakao_window_class,
            kakao_main_title=kakao_main_title,
            kakao_room_name=kakao_room_name,
            message_template=message_template,
            kakao_chat_tab_x=kakao_chat_tab_x,
            kakao_chat_tab_y=kakao_chat_tab_y,
            kakao_input_x=kakao_input_x,
            kakao_input_y=kakao_input_y,
            kakao_room_window_wait_seconds=kakao_room_window_wait_seconds,
            kakao_foreground_retry_pause_seconds=kakao_foreground_retry_pause_seconds,
            kakao_window_poll_interval_seconds=kakao_window_poll_interval_seconds,
            kakao_activate_show_pause_seconds=kakao_activate_show_pause_seconds,
            kakao_after_activate_pause_seconds=kakao_after_activate_pause_seconds,
            kakao_chat_tab_pause_seconds=kakao_chat_tab_pause_seconds,
            kakao_search_open_pause_seconds=kakao_search_open_pause_seconds,
            kakao_search_reset_pause_seconds=kakao_search_reset_pause_seconds,
            kakao_search_paste_pause_seconds=kakao_search_paste_pause_seconds,
            kakao_room_enter_pause_seconds=kakao_room_enter_pause_seconds,
            kakao_input_click_pause_seconds=kakao_input_click_pause_seconds,
            kakao_paste_pause_seconds=kakao_paste_pause_seconds,
            kakao_send_pause_seconds=kakao_send_pause_seconds,
            kakao_search_clear_backspace_count=kakao_search_clear_backspace_count,
            stop_flag_path=stop_flag_path,
            log_level=log_level,
            log_path=log_path,
            config_path=path,
        )
