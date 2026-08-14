from __future__ import annotations

from pathlib import Path

import pytest

from config import (
    FORESTBOND_TITLE,
    KBOND_PROCESS,
    NOTEPAD_PROCESS,
    NOTEPAD_TITLE,
    Config,
    ConfigError,
)
from source import KbondSourceReader, UiaSourceReader, create_source_reader


def _write_env(path: Path, mode: str, extra: str = "") -> Path:
    path.write_text(
        "\n".join(
            [
                f"MODE={mode}",
                "POLL_INTERVAL_MS=300",
                "PROCESS_EXISTING_ON_START=true",
                "EXCEL_WORKBOOK=sample.xlsm",
                "EXCEL_SHEET=트레이딩",
                "EXCEL_SLOT_ROWS=19,25,41,46,56",
                "EXCEL_ROWS_10Y=19,25",
                "EXCEL_ROWS_3Y=41,46,56",
                "EXCEL_INSTRUMENT_COL=A",
                "EXCEL_QTY_COL=E",
                "EXCEL_INPUT_COL=D",
                "EXCEL_PNL_COL=F",
                "EXCEL_PNL_ROW_OFFSET=3",
                "EXCEL_PREFIX_3Y_CELL=B5",
                "EXCEL_PREFIX_10Y_CELL=B6",
                "EXCEL_WATCH_CELL=D2",
                "EXCEL_PNL_THRESHOLD_CELL=E2",
                "EXCEL_PNL_SANITY_BAND=5000000",
                "EXCEL_STATUS_CELL=F2",
                "EXCEL_LOOKING_FOR_CELL=G2",
                "EXCEL_LAST_QUOTE_CELL=H2",
                "EXCEL_LAST_PNL_CELL=I2",
                "EXCEL_LAST_ACTION_CELL=J2",
                "MESSAGE_TEMPLATE={instrument} {confirm_token} ㅎㅈ",
                "KBOND_CHAT_TITLE=[채권] 블커본드",
                "SEND_INPUT_X_M1=0.825",
                "SEND_INPUT_Y_M1=0.940",
                "SEND_INPUT_X_M23=0.5",
                "SEND_INPUT_Y_M23=0.5",
                "SEND_FOREGROUND_RETRY_PAUSE_SECONDS=0.05",
                "SEND_ACTIVATE_SHOW_PAUSE_SECONDS=0.05",
                "SEND_AFTER_ACTIVATE_PAUSE_SECONDS=0.10",
                "SEND_INPUT_CLICK_PAUSE_SECONDS=0.08",
                "SEND_PASTE_PAUSE_SECONDS=0.10",
                "SEND_SEND_PAUSE_SECONDS=0.08",
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


CHAT_TITLE = "[채권] 블커본드"


def test_mode1_presets(tmp_path: Path) -> None:
    cfg = Config.load(_write_env(tmp_path / "m1.env", "1"))
    assert cfg.mode == 1
    assert cfg.source_window_title == CHAT_TITLE
    assert cfg.source_process_name == KBOND_PROCESS
    assert cfg.send_process_name == KBOND_PROCESS
    assert cfg.send_window_title == CHAT_TITLE
    assert cfg.send_input_x == pytest.approx(0.825)
    assert cfg.send_input_y == pytest.approx(0.940)
    assert isinstance(create_source_reader(cfg), KbondSourceReader)


def test_mode2_presets(tmp_path: Path) -> None:
    cfg = Config.load(_write_env(tmp_path / "m2.env", "2"))
    assert cfg.mode == 2
    assert cfg.source_window_title == CHAT_TITLE
    assert cfg.source_process_name == KBOND_PROCESS
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert cfg.send_window_title == NOTEPAD_TITLE
    assert cfg.send_input_x == pytest.approx(0.5)
    assert cfg.send_input_y == pytest.approx(0.5)
    assert cfg.excel_watch_cell == "D2"
    assert cfg.excel_pnl_threshold_cell == "E2"
    assert cfg.excel_pnl_sanity_band == pytest.approx(5_000_000)
    assert cfg.send_foreground_retry_pause_seconds == pytest.approx(0.05)
    assert cfg.send_activate_show_pause_seconds == pytest.approx(0.05)
    assert cfg.send_after_activate_pause_seconds == pytest.approx(0.10)
    assert cfg.send_input_click_pause_seconds == pytest.approx(0.08)
    assert cfg.send_paste_pause_seconds == pytest.approx(0.10)
    assert cfg.send_send_pause_seconds == pytest.approx(0.08)
    assert isinstance(create_source_reader(cfg), KbondSourceReader)


def test_mode3_presets(tmp_path: Path) -> None:
    cfg = Config.load(_write_env(tmp_path / "m3.env", "3"))
    assert cfg.mode == 3
    assert cfg.source_window_title == FORESTBOND_TITLE
    assert cfg.source_process_name == ""
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert cfg.send_window_title == NOTEPAD_TITLE
    assert cfg.send_input_x == pytest.approx(0.5)
    assert cfg.send_input_y == pytest.approx(0.5)
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
    assert cfg.source_window_title == CHAT_TITLE
    assert cfg.send_process_name == NOTEPAD_PROCESS
    assert cfg.send_input_x == pytest.approx(0.5)


def test_mode1_uses_m1_ratios_from_env(tmp_path: Path) -> None:
    cfg = Config.load(
        _write_env(
            tmp_path / "m1_ratio.env",
            "1",
            extra="SEND_INPUT_X_M1=0.71\nSEND_INPUT_Y_M1=0.92\n",
        )
    )
    assert cfg.send_input_x == pytest.approx(0.71)
    assert cfg.send_input_y == pytest.approx(0.92)


def test_mode23_uses_m23_ratios_from_env(tmp_path: Path) -> None:
    cfg = Config.load(
        _write_env(
            tmp_path / "m3_ratio.env",
            "3",
            extra="SEND_INPUT_X_M23=0.4\nSEND_INPUT_Y_M23=0.6\n",
        )
    )
    assert cfg.send_input_x == pytest.approx(0.4)
    assert cfg.send_input_y == pytest.approx(0.6)


@pytest.mark.parametrize("bad", ["0", "4", "-1", "abc"])
def test_invalid_mode(tmp_path: Path, bad: str) -> None:
    with pytest.raises(ConfigError):
        Config.load(_write_env(tmp_path / f"bad_{bad}.env", bad))


def test_empty_watch_cells_rejected(tmp_path: Path) -> None:
    with pytest.raises(ConfigError, match="EXCEL_WATCH_CELL"):
        Config.load(
            _write_env(tmp_path / "empty_watch.env", "2", extra="EXCEL_WATCH_CELL=\n")
        )
    with pytest.raises(ConfigError, match="EXCEL_PNL_THRESHOLD_CELL"):
        Config.load(
            _write_env(
                tmp_path / "empty_thr.env",
                "2",
                extra="EXCEL_PNL_THRESHOLD_CELL=\n",
            )
        )


def test_mode12_requires_kbond_chat_title(tmp_path: Path) -> None:
    for mode in ("1", "2"):
        with pytest.raises(ConfigError, match="KBOND_CHAT_TITLE"):
            Config.load(
                _write_env(
                    tmp_path / f"no_chat_{mode}.env",
                    mode,
                    extra="KBOND_CHAT_TITLE=\n",
                )
            )


def test_mode3_loads_without_kbond_chat_title(tmp_path: Path) -> None:
    path = _write_env(tmp_path / "m3_no_chat.env", "3")
    lines = [
        line
        for line in path.read_text(encoding="utf-8").splitlines()
        if not line.startswith("KBOND_CHAT_TITLE=")
    ]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")
    cfg = Config.load(path)
    assert cfg.mode == 3
    assert cfg.source_window_title == FORESTBOND_TITLE


def test_sanity_band_must_be_positive(tmp_path: Path) -> None:
    with pytest.raises(ConfigError, match="EXCEL_PNL_SANITY_BAND"):
        Config.load(
            _write_env(
                tmp_path / "bad_band.env",
                "2",
                extra="EXCEL_PNL_SANITY_BAND=0\n",
            )
        )
