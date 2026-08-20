from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path


class ConfigError(ValueError):
    pass


KBOND_PROCESS = "KBondMessenger.exe"
KBOND_TITLE = "K-Bond"
FORESTBOND_TITLE = "FORESTBOND"
NOTEPAD_PROCESS = "notepad.exe"
NOTEPAD_TITLE = "메모장"


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
    excel_pnl_sanity_band: float

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
