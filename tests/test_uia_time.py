from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from source.reader_uia import attach_preceding_time, UiaSourceReader, watermark_has_clock


def test_attach_preceding_time_binds_immediate_time_line() -> None:
    items = attach_preceding_time(
        [
            "권** (17:48:01) :",
            "26-3 005 01 + 20",
            "(**증권 종합금융팀 02-****)",
        ]
    )
    assert items[0].text == "권** (17:48:01) :"
    assert items[0].watermark_key == "권** (17:48:01) :"
    assert items[1].text == "26-3 005 01 + 20"
    assert items[1].watermark_key == "(17:48:01) : 26-3 005 01 + 20"
    assert items[2].text == "(**증권 종합금융팀 02-****)"
    assert items[2].watermark_key == "(**증권 종합금융팀 02-****)"


def test_attach_preceding_time_bare_clock() -> None:
    items = attach_preceding_time(["(17:48:01) :", "25-10 695 +"])
    assert items[1].text == "25-10 695 +"
    assert items[1].watermark_key == "(17:48:01) : 25-10 695 +"


def test_attach_preceding_time_no_time_keeps_quote() -> None:
    items = attach_preceding_time(["26-3 005 01 + 20"])
    assert items[0].text == "26-3 005 01 + 20"
    assert items[0].watermark_key == "26-3 005 01 + 20"


def test_attach_preceding_time_does_not_split_combined_line() -> None:
    full = "권** (17:48:01) : 26-3 005 01 + 20"
    items = attach_preceding_time([full])
    assert len(items) == 1
    assert items[0].text == full
    assert items[0].watermark_key == full


def test_attach_preceding_time_same_quote_two_clocks() -> None:
    items = attach_preceding_time(
        [
            "(17:48:01) :",
            "26-3 005 01 + 20",
            "(17:48:02) :",
            "26-3 005 01 + 20",
        ]
    )
    quotes = [item for item in items if item.text == "26-3 005 01 + 20"]
    assert [item.watermark_key for item in quotes] == [
        "(17:48:01) : 26-3 005 01 + 20",
        "(17:48:02) : 26-3 005 01 + 20",
    ]


def test_watermark_has_clock() -> None:
    assert watermark_has_clock("(17:48:01) : 25-10 695 +")
    assert watermark_has_clock("권** (17:48) :")
    assert not watermark_has_clock("26-3 005 01 + 20")
    assert not watermark_has_clock("(**증권 종합금융팀 02-****)")


def test_uia_reuses_window_when_handle_valid(monkeypatch: pytest.MonkeyPatch) -> None:
    reader = UiaSourceReader("FORESTBOND")
    calls = {"n": 0}

    class _Ctrl:
        def window_text(self) -> str:
            return "hello"

    class _Win:
        handle = 42

        def descendants(self, control_type: str | None = None) -> list[_Ctrl]:
            return [_Ctrl()]

    def _find() -> _Win:
        calls["n"] += 1
        win = _Win()
        reader._window = win
        return win

    monkeypatch.setattr(reader, "find_source_window", _find)
    monkeypatch.setattr("source.reader_uia.win32gui.IsWindow", lambda hwnd: True)
    monkeypatch.setattr("source.reader_uia.win32gui.IsWindowVisible", lambda hwnd: True)

    reader.get_visible_message_lines()
    reader.get_visible_message_lines()
    assert calls["n"] == 1


def test_uia_text_enum_retries_then_succeeds(monkeypatch: pytest.MonkeyPatch) -> None:
    reader = UiaSourceReader("FORESTBOND")
    calls = {"n": 0}

    class _Ctrl:
        def window_text(self) -> str:
            return "hello"

    class _Win:
        handle = 42

        def descendants(self, control_type: str | None = None) -> list[_Ctrl]:
            calls["n"] += 1
            if calls["n"] == 1:
                raise RuntimeError("transient")
            return [_Ctrl()]

    monkeypatch.setattr(reader, "_ensure_window", lambda: _Win())
    monkeypatch.setattr("source.reader_uia.time.sleep", lambda s: None)
    lines = reader.get_visible_message_lines()
    assert calls["n"] == 2
    assert lines[0].text == "hello"


def test_uia_text_enum_retries_then_errors(monkeypatch: pytest.MonkeyPatch) -> None:
    from source.common import SourceReaderError

    reader = UiaSourceReader("FORESTBOND")
    calls = {"n": 0}

    class _Win:
        handle = 42

        def descendants(self, control_type: str | None = None) -> list:
            calls["n"] += 1
            raise RuntimeError("always")

    monkeypatch.setattr(reader, "_ensure_window", lambda: _Win())
    monkeypatch.setattr("source.reader_uia.time.sleep", lambda s: None)
    with pytest.raises(SourceReaderError, match="Failed to enumerate Text controls"):
        reader.get_visible_message_lines()
    assert calls["n"] == 3
