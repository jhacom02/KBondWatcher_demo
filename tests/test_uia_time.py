from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from source.common import SourceReaderError, source_line
from source.reader_uia import UiaSourceReader, new_lines_after


def _keys(lines: list) -> list[str]:
    return [item.watermark_key for item in lines]


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


def test_new_lines_after_empty_prev() -> None:
    assert new_lines_after([], [source_line("a")]) == []


def test_new_lines_after_same_snapshot() -> None:
    snap = [source_line("x"), source_line("y")]
    assert new_lines_after(snap, snap) == []


def test_new_lines_after_append() -> None:
    prev = [source_line("a"), source_line("b")]
    now = [source_line("a"), source_line("b"), source_line("c")]
    assert _keys(new_lines_after(prev, now)) == ["c"]


def test_new_lines_after_prefix_drop_then_append() -> None:
    prev = [
        source_line("chrome"),
        source_line("old1"),
        source_line("old2"),
        source_line("quote"),
    ]
    now = [
        source_line("chrome"),
        source_line("old2"),
        source_line("quote"),
        source_line("new"),
    ]
    assert _keys(new_lines_after(prev, now)) == ["new"]


def test_new_lines_after_no_overlap() -> None:
    prev = [source_line("a"), source_line("b")]
    now = [source_line("x"), source_line("y")]
    assert new_lines_after(prev, now) == []


def test_new_lines_after_prepend_after_chrome() -> None:
    prev = [
        source_line("chrome"),
        source_line("25-10 76+"),
        source_line("국주"),
    ]
    now = [
        source_line("chrome"),
        source_line("25-10 755-"),
        source_line("25-10 76+"),
        source_line("국주"),
    ]
    assert _keys(new_lines_after(prev, now)) == ["25-10 755-"]


def test_new_lines_after_in_place_replace() -> None:
    prev = [source_line("chrome"), source_line("25-10 76+")]
    now = [source_line("chrome"), source_line("25-10 77+")]
    assert _keys(new_lines_after(prev, now)) == ["25-10 77+"]


def test_new_lines_after_duplicate_string_appended() -> None:
    prev = [source_line("chrome"), source_line("25-10 77+")]
    now = [
        source_line("chrome"),
        source_line("25-10 77+"),
        source_line("25-10 77+"),
    ]
    assert _keys(new_lines_after(prev, now)) == ["25-10 77+"]


def test_new_lines_after_reorder_same_counts_is_empty() -> None:
    prev = [
        source_line("chrome"),
        source_line("25-10 76+"),
        source_line("국주"),
    ]
    now = [
        source_line("chrome"),
        source_line("국주"),
        source_line("25-10 76+"),
    ]
    assert new_lines_after(prev, now) == []


def test_uia_visible_keeps_duplicate_quote_nodes(monkeypatch: pytest.MonkeyPatch) -> None:
    reader = UiaSourceReader("FORESTBOND")

    class _Ctrl:
        def __init__(self, text: str) -> None:
            self._text = text

        def window_text(self) -> str:
            return self._text

    class _Win:
        handle = 42

        def descendants(self, control_type: str | None = None) -> list[_Ctrl]:
            return [_Ctrl("25-10 77+"), _Ctrl("25-10 77+")]

    monkeypatch.setattr(reader, "_ensure_window", lambda: _Win())
    lines = reader.get_visible_message_lines()
    assert [item.text for item in lines] == ["25-10 77+", "25-10 77+"]


def test_uia_get_new_message_lines_same_snapshot_is_empty() -> None:
    reader = UiaSourceReader("FORESTBOND")
    quote = source_line("25-10 775+")
    reader.get_visible_message_lines = lambda: [quote]  # type: ignore[method-assign]
    reader.initialize_watermark(False)
    first = reader.get_new_message_lines()
    second = reader.get_new_message_lines()
    assert first == []
    assert second == []


def test_uia_get_new_message_lines_returns_suffix() -> None:
    reader = UiaSourceReader("FORESTBOND")
    snapshots: list[list] = [
        [source_line("a"), source_line("b")],
        [source_line("a"), source_line("b")],
        [source_line("a"), source_line("b"), source_line("c")],
    ]
    state = {"n": 0}

    def _visible() -> list:
        idx = min(state["n"], len(snapshots) - 1)
        state["n"] += 1
        return list(snapshots[idx])

    reader.get_visible_message_lines = _visible  # type: ignore[method-assign]
    reader.initialize_watermark(False)
    assert reader.get_new_message_lines() == []
    got = reader.get_new_message_lines()
    assert [item.text for item in got] == ["c"]
