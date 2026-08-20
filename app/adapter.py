from __future__ import annotations

from dataclasses import replace
from pathlib import Path

from config import Config, mode_presets
from excel.bridge import InstrumentSlot

from .defaults import DEFAULTS, DeveloperDefaults
from .machine import MachineCalibration
from .profile import TraderProfile
from .side_map import normalize_looking_for, required_side_from_looking_for


def slot_from_profile(profile: TraderProfile, yield_prefix: float) -> InstrumentSlot:
    looking = normalize_looking_for(profile.looking_for)
    return InstrumentSlot(
        instrument=profile.instrument.strip(),
        row=0,
        looking_for=looking,
        required_side=required_side_from_looking_for(looking),
        qty_abs=int(profile.required_qty),
        yield_prefix=float(yield_prefix),
        input_cell=profile.yield_input_cell.strip().upper(),
        qty_cell="",
        pnl_cell=profile.pnl_cell.strip().upper(),
    )


def config_from_profile(
    profile: TraderProfile,
    machine: MachineCalibration,
    defaults: DeveloperDefaults = DEFAULTS,
    *,
    config_path: Path | None = None,
) -> Config:
    mode = int(profile.mode)
    source_window_title, source_process_name, send_process_name, send_window_title = (
        mode_presets(mode)
    )
    if mode in (1, 2):
        source_window_title = profile.kbond_chat_title.strip()
        if mode == 1:
            send_window_title = source_window_title

    return Config(
        mode=mode,
        source_window_title=source_window_title,
        source_process_name=source_process_name,
        poll_interval_ms=defaults.poll_interval_ms,
        process_existing_on_start=defaults.process_existing_on_start,
        excel_workbook=profile.excel_workbook.strip(),
        excel_sheet=profile.excel_sheet.strip(),
        excel_pnl_sanity_band=defaults.excel_pnl_sanity_band,
        send_process_name=send_process_name,
        send_window_title=send_window_title,
        message_template=profile.message_template,
        send_input_x=float(machine.send_input_x),
        send_input_y=float(machine.send_input_y),
        send_foreground_retry_pause_seconds=defaults.send_foreground_retry_pause_seconds,
        send_activate_show_pause_seconds=defaults.send_activate_show_pause_seconds,
        send_after_activate_pause_seconds=defaults.send_after_activate_pause_seconds,
        send_input_click_pause_seconds=defaults.send_input_click_pause_seconds,
        send_paste_pause_seconds=defaults.send_paste_pause_seconds,
        send_send_pause_seconds=defaults.send_send_pause_seconds,
        stop_flag_path=defaults.stop_flag,
        log_level=defaults.log_level,
        log_path=defaults.log_path,
        config_path=config_path or Path("profile"),
        sent_after=(profile.sent_after or defaults.sent_after_default).strip().lower(),
    )


def with_send_ratios(cfg: Config, x: float, y: float) -> Config:
    return replace(cfg, send_input_x=x, send_input_y=y)
