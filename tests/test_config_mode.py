from __future__ import annotations

from pathlib import Path

import pytest

from config import (
    FORESTBOND_TITLE,
    KBOND_PROCESS,
    KBOND_SEND_INPUT_X,
    KBOND_SEND_INPUT_Y,
    KBOND_TITLE,
    NOTEPAD_PROCESS,
    NOTEPAD_SEND_INPUT_X,
    NOTEPAD_SEND_INPUT_Y,
    NOTEPAD_TITLE,
    Config,
    ConfigError,
)
from source_reader import create_source_reader
from source_reader_kbond import KbondSourceReader
from source_reader_uia import UiaSourceReader


def _write_env(path: Path, mode: str, extra: str = "") -> Path:
    path.write_text(
        "\n".join(
            [
                f"MODE={mode}",
                "POLL_INTERVAL_MS=300",
                "PROCESS_EXISTING_ON_START=true",
                "PNL_THRESHOLD=1000000",
                "EXCEL_WORKBOOK=sample.xlsm",
                "EXCEL_SHEET=트레이딩",
                "EXCEL_SLOT_ROWS=19,25,41,46,56",
                "EXCEL_ROWS_10Y=19,25",
                "EXCEL_ROWS_3Y=41,46,56",
                "EXCEL_PREFIX_3Y_CELL=B5",
                "EXCEL_PREFIX_10Y_CELL=B6",
                "EXCEL_STATUS_CELL=F2",
                "EXCEL_LOOKING_FOR_CELL=G2",
                "EXCEL_LAST_QUOTE_CELL=H2",
                "EXCEL_LAST_PNL_CELL=I2",
                "EXCEL_LAST_ACTION_CELL=J2",
                "MESSAGE_TEMPLATE={instrument} {confirm_token} ㅎㅈ",
                "SEND_FOREGROUND_RETRY_PAUSE_SECONDS=0.1",
                "SEND_ACTIVATE_SHOW_PAUSE_SECONDS=0.15",
                "SEND_AFTER_ACTIVATE_PAUSE_SECONDS=0.2",
                "SEND_INPUT_CLICK_PAUSE_SECONDS=0.2",
                "SEND_PASTE_PAUSE_SECONDS=0.3",
                "SEND_SEND_PAUSE_SECONDS=0.3",
                "STOP_FLAG_PATH=C:\\temp\\kbond_watcher.stop",
                "LOG_LEVEL=INFO",
                "LOG_PATH=logs\\watcher.log",
                extra,
            ]
        ).strip()
        + "\n",
        encoding="utf-8",
    )
    return path


def test_mode1_presets(tmp_path: Path) -> None:
    cfg = Config.load(_write_env(tmp_path / "m1.env", "1"))
    assert cfg.mode == 1
    assert cfg.source_window_title == KBOND_TITLE
    assert cfg.source_process_name == KBOND_PROCESS
    assert cfg.send_process_name == KBOND_PROCESS
    assert cfg.send_window_title == KBOND_TITLE
    assert cfg.send_input_x == KBOND_SEND_INPUT_X
    assert cfg.send_input_y == KBOND_SEND_INPUT_Y
    assert isinstance(create_source_reader(cfg), KbondSourceReader)


def test_mode2_presets(tmp_path: Path) -> None:
    cfg = Config.load(_write_env(tmp_path / "m2.env", "2"))
    assert cfg.mode == 2
    assert cfg.source_window_title == KBOND_TITLE
    assert cfg.source_process_name == KBOND_PROCESS
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert cfg.send_window_title == NOTEPAD_TITLE
    assert cfg.send_input_x == NOTEPAD_SEND_INPUT_X
    assert cfg.send_input_y == NOTEPAD_SEND_INPUT_Y
    assert isinstance(create_source_reader(cfg), KbondSourceReader)


def test_mode3_presets(tmp_path: Path) -> None:
    cfg = Config.load(_write_env(tmp_path / "m3.env", "3"))
    assert cfg.mode == 3
    assert cfg.source_window_title == FORESTBOND_TITLE
    assert cfg.source_process_name == ""
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert cfg.send_window_title == NOTEPAD_TITLE
    assert cfg.send_input_x == NOTEPAD_SEND_INPUT_X
    assert cfg.send_input_y == NOTEPAD_SEND_INPUT_Y
    assert isinstance(create_source_reader(cfg), UiaSourceReader)


def test_mode_ignores_conflicting_identity_keys(tmp_path: Path) -> None:
    cfg = Config.load(
        _write_env(
            tmp_path / "conflict.env",
            "2",
            extra=(
                "SOURCE_WINDOW_TITLE=WRONG\n"
                "SOURCE_PROCESS_NAME=wrong.exe\n"
                "SEND_PROCESS_NAME=wrong.exe\n"
                "SEND_WINDOW_TITLE=WRONG\n"
                "SEND_INPUT_X=0.1\n"
                "SEND_INPUT_Y=0.1\n"
            ),
        )
    )
    assert cfg.source_window_title == KBOND_TITLE
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert cfg.send_input_x == NOTEPAD_SEND_INPUT_X


@pytest.mark.parametrize("bad", ["0", "4", "-1", "abc"])
def test_invalid_mode(tmp_path: Path, bad: str) -> None:
    with pytest.raises(ConfigError):
        Config.load(_write_env(tmp_path / f"bad_{bad}.env", bad))
