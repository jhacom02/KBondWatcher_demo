from __future__ import annotations

import pytest

from app.adapter import config_from_profile
from app.defaults import DEFAULTS
from app.machine import MachineCalibration
from app.profile import ProfileError, TraderProfile
from config import (
    FORESTBOND_TITLE,
    KBOND_PROCESS,
    NOTEPAD_PROCESS,
    NOTEPAD_TITLE,
    ConfigError,
    mode_presets,
)
from source import KbondSourceReader, UiaSourceReader, create_source_reader


def _profile(**overrides) -> TraderProfile:
    data = dict(
        profile_name="t1",
        profile_version=1,
        trader_id="t1",
        instrument="25-10",
        looking_for="BID",
        required_qty=100,
        threshold=-80.0,
        threshold_op="<=",
        excel_workbook=r"C:\Trading\bond.xlsm",
        excel_sheet="Sheet1",
        yield_input_cell="D41",
        pnl_cell="F44",
        yield_prefix=3.0,
        mode=2,
        kbond_chat_title="[채권] 블커본드",
        sent_after="exit",
        message_template="{instrument} {confirm_token} ㅎㅈ",
    )
    data.update(overrides)
    return TraderProfile(**data)


def _machine(**overrides) -> MachineCalibration:
    data = dict(machine_id="m1", send_input_x=0.5, send_input_y=0.9)
    data.update(overrides)
    return MachineCalibration(**data)


def test_mode_presets() -> None:
    assert mode_presets(1)[1] == KBOND_PROCESS
    assert mode_presets(2)[2:] == (NOTEPAD_PROCESS, NOTEPAD_TITLE)
    assert mode_presets(3)[0] == FORESTBOND_TITLE
    with pytest.raises(ConfigError, match="MODE must be 1, 2, or 3"):
        mode_presets(9)


def test_config_from_profile_mode1() -> None:
    cfg = config_from_profile(
        _profile(mode=1, kbond_chat_title="[채권] A"),
        _machine(send_input_x=0.2, send_input_y=0.3),
    )
    assert cfg.mode == 1
    assert cfg.source_window_title == "[채권] A"
    assert cfg.send_window_title == "[채권] A"
    assert cfg.source_process_name == KBOND_PROCESS
    assert cfg.send_process_name == KBOND_PROCESS
    assert cfg.send_input_x == pytest.approx(0.2)
    assert cfg.send_input_y == pytest.approx(0.3)
    assert cfg.excel_pnl_sanity_band == pytest.approx(DEFAULTS.excel_pnl_sanity_band)
    assert cfg.sent_after == "exit"
    assert isinstance(create_source_reader(cfg), KbondSourceReader)


def test_config_from_profile_mode2() -> None:
    cfg = config_from_profile(_profile(mode=2), _machine())
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert cfg.send_window_title == NOTEPAD_TITLE
    assert cfg.source_window_title == "[채권] 블커본드"
    assert isinstance(create_source_reader(cfg), KbondSourceReader)


def test_config_from_profile_mode3() -> None:
    cfg = config_from_profile(_profile(mode=3, kbond_chat_title=""), _machine())
    assert cfg.source_window_title == FORESTBOND_TITLE
    assert cfg.source_process_name == ""
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert isinstance(create_source_reader(cfg), UiaSourceReader)


def test_config_from_profile_loop() -> None:
    cfg = config_from_profile(_profile(mode=2, sent_after="loop"), _machine())
    assert cfg.sent_after == "loop"


def test_profile_mode1_rejects_loop() -> None:
    with pytest.raises(ProfileError, match="sent_after must be exit when mode is 1"):
        _profile(mode=1, sent_after="loop").validate()


def test_profile_rejects_bad_sent_after() -> None:
    with pytest.raises(ProfileError, match="sent_after must be exit or loop"):
        _profile(sent_after="maybe").validate()
