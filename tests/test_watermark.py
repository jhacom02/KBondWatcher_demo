from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from source.common import (
    BaseSourceReader,
    SourceLine,
    WATERMARK_WINDOW,
    as_source_lines,
    message_fingerprint,
    source_line,
    watermark_window,
)


class FakeReader(BaseSourceReader):
    def __init__(self, lines: list[str] | None = None) -> None:
        super().__init__()
        self.lines = list(lines or [])

    def find_source_window(self) -> int:
        return 1

    def get_visible_message_lines(self) -> list[SourceLine]:
        return as_source_lines(self.lines)

    def diagnose(self, max_messages: int = 200) -> str:
        return ""


def _texts(items: list[SourceLine]) -> list[str]:
    return [item.text for item in items]


def test_watermark_window_returns_tail() -> None:
    lines = [f"L{i}" for i in range(5)]
    assert watermark_window(lines, window=3) == ["L2", "L3", "L4"]
    assert watermark_window(lines, window=10) == lines


def test_init_false_skips_existing_then_emits_tail_only(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr("source.common.WATERMARK_WINDOW", 3)
    reader = FakeReader(["A", "B", "C", "D"])
    reader.initialize_watermark(False)
    assert _texts(reader.get_new_message_lines()) == []
    reader.lines = ["A", "B", "C", "D", "E"]
    assert _texts(reader.get_new_message_lines()) == ["E"]


def test_head_only_change_is_ignored(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("source.common.WATERMARK_WINDOW", 3)
    reader = FakeReader(["B", "C", "D"])
    reader.initialize_watermark(False)
    reader.lines = ["OLD", "B", "C", "D"]
    assert _texts(reader.get_new_message_lines()) == []


def test_fifo_evicts_oldest_and_does_not_resurrect_head(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr("source.common.WATERMARK_WINDOW", 3)
    reader = FakeReader(["A", "B", "C"])
    reader.initialize_watermark(False)
    reader.lines = ["B", "C", "D"]
    assert _texts(reader.get_new_message_lines()) == ["D"]
    assert len(reader._watermark) == 3
    assert message_fingerprint("A") not in reader._watermark
    reader.lines = ["A", "B", "C", "D"]
    assert _texts(reader.get_new_message_lines()) == []


def test_process_existing_true_returns_window(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    monkeypatch.setattr("source.common.WATERMARK_WINDOW", 3)
    reader = FakeReader(["A", "B", "C", "D"])
    assert _texts(
        reader.get_new_message_lines(process_existing_on_start=True)
    ) == ["B", "C", "D"]


def test_reseed_unions_window_only(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("source.common.WATERMARK_WINDOW", 2)
    reader = FakeReader(["A", "B"])
    reader.initialize_watermark(False)
    reader.lines = ["A", "B", "C"]
    reader.reseed_watermark_from_visible()
    assert _texts(reader.get_new_message_lines()) == []
    reader.lines = ["B", "C", "D"]
    assert _texts(reader.get_new_message_lines()) == ["D"]


def test_fifo_caps_set_size(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("source.common.WATERMARK_WINDOW", 3)
    reader = FakeReader(["L0", "L1", "L2"])
    reader.initialize_watermark(False)
    for i in range(3, 10):
        reader.lines = [f"L{i - 2}", f"L{i - 1}", f"L{i}"]
        new = reader.get_new_message_lines()
        assert _texts(new) == [f"L{i}"]
        assert len(reader._watermark) == 3
        assert len(reader._watermark_order) == 3


def test_watermark_uses_key_not_text(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setattr("source.common.WATERMARK_WINDOW", 5)

    class KeyReader(FakeReader):
        def get_visible_message_lines(self) -> list[SourceLine]:
            return [
                source_line(text, key)
                for text, key in [
                    ("26-3 005 01 + 20", "(17:48:01) : 26-3 005 01 + 20"),
                    ("26-3 005 01 + 20", "(17:48:02) : 26-3 005 01 + 20"),
                ]
            ]

    reader = KeyReader()
    got = reader.get_new_message_lines(process_existing_on_start=True)
    assert _texts(got) == ["26-3 005 01 + 20", "26-3 005 01 + 20"]
    assert got[0].watermark_key != got[1].watermark_key


def test_default_window_constant() -> None:
    assert WATERMARK_WINDOW == 2000
