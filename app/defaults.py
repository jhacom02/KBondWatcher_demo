from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from .paths import logs_dir, stop_flag_path


@dataclass(frozen=True)
class DeveloperDefaults:
    poll_interval_ms: int = 300
    process_existing_on_start: bool = False
    excel_pnl_sanity_band: float = 5_000_000.0
    send_foreground_retry_pause_seconds: float = 0.05
    send_activate_show_pause_seconds: float = 0.05
    send_after_activate_pause_seconds: float = 0.05
    send_input_click_pause_seconds: float = 0.05
    send_paste_pause_seconds: float = 0.05
    send_send_pause_seconds: float = 0.05
    stop_soft_wait_seconds: float = 8.0
    heartbeat_stale_seconds: float = 5.0
    lease_refresh_seconds: float = 60.0
    policy_poll_seconds: float = 60.0
    perf_pad_seconds: float = 0.5
    log_level: str = "INFO"
    sent_after_default: str = "exit"

    @property
    def stop_flag(self) -> Path:
        return stop_flag_path()

    @property
    def log_path(self) -> Path:
        return logs_dir() / "watcher.log"


DEFAULTS = DeveloperDefaults()
