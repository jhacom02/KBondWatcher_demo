from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import pytest

from source.eltree import normalize_lines
from source.reader_kbond import KbondSourceReader
from source.richedit import (
    MAX_TEXT_CHARS,
    EditCandidate,
    RichEditSnapshot,
    is_gettext_clipped,
    pick_chat_edit,
    resolve_gettext_read,
)


def test_pick_chat_edit_prefers_largest_visible_pane() -> None:
    tiny = EditCandidate(1, 10, "a", "TfrmDetach", 0, 0, 50, 20, True)
    inputish = EditCandidate(2, 10, "a", "TfrmDetach", 0, 900, 400, 950, True)
    chat = EditCandidate(3, 10, "[채팅] room", "TfrmDetach", 0, 0, 470, 750, True)
    chosen = pick_chat_edit([tiny, inputish, chat])
    assert chosen is not None
    assert chosen.hwnd == 3


def test_pick_chat_edit_rejects_small_only() -> None:
    tiny = EditCandidate(1, 10, "a", "TfrmDetach", 0, 0, 80, 40, True)
    assert pick_chat_edit([tiny]) is None


def test_normalize_richedit_text_splits_crlf() -> None:
    raw = "입장하셨습니다.\r\n홍길동 (15:39:36) : 26-6 68+\r\n입장하셨습니다.\r\n"
    lines = normalize_lines([raw])
    assert lines[0] == "입장하셨습니다."
    assert "26-6 68+" in lines[1]
    assert lines.count("입장하셨습니다.") == 1


def test_is_gettext_clipped_at_cap() -> None:
    assert not is_gettext_clipped(MAX_TEXT_CHARS)
    assert is_gettext_clipped(MAX_TEXT_CHARS + 1)


def test_resolve_gettext_read_uses_cache_when_len_unchanged() -> None:
    assert resolve_gettext_read(100, 100) == "use_cache"
    assert resolve_gettext_read(101, 100) == "fetch"
    assert resolve_gettext_read(50, None) == "fetch"


def test_resolve_gettext_read_skips_when_clipped() -> None:
    assert resolve_gettext_read(MAX_TEXT_CHARS + 1, 100) == "skip_clip"
    assert resolve_gettext_read(MAX_TEXT_CHARS + 1, None) == "skip_clip"


def test_richedit_cache_skips_snapshot_when_len_unchanged(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    reader = KbondSourceReader("K-Bond", "KBondMessenger.exe")
    reader._backend = "richedit"
    reader._edit_hwnd = 1
    reader._cache_len = 10
    reader._cached_lines = ["a", "b"]
    snap_calls = {"n": 0}

    monkeypatch.setattr("source.reader_kbond.read_richedit_length", lambda hwnd: 10)

    def _snap(hwnd: int, max_chars: int = 0) -> RichEditSnapshot:
        snap_calls["n"] += 1
        return RichEditSnapshot(lines=["x"], char_len=10, clipped=False)

    monkeypatch.setattr("source.reader_kbond.read_richedit_snapshot", _snap)
    assert reader._read_richedit_lines() == ["a", "b"]
    assert snap_calls["n"] == 0


def test_richedit_clip_keeps_last_good_cache(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    reader = KbondSourceReader("K-Bond", "KBondMessenger.exe")
    reader._backend = "richedit"
    reader._edit_hwnd = 1
    reader._cache_len = 10
    reader._cached_lines = ["good"]
    snap_calls = {"n": 0}

    monkeypatch.setattr(
        "source.reader_kbond.read_richedit_length",
        lambda hwnd: MAX_TEXT_CHARS + 5,
    )

    def _snap(hwnd: int, max_chars: int = 0) -> RichEditSnapshot:
        snap_calls["n"] += 1
        return RichEditSnapshot(lines=["truncated-head"], char_len=1, clipped=True)

    monkeypatch.setattr("source.reader_kbond.read_richedit_snapshot", _snap)
    assert reader._read_richedit_lines() == ["good"]
    assert reader._last_clipped is True
    assert reader._last_gettext_len == MAX_TEXT_CHARS + 5
    assert snap_calls["n"] == 0


def test_diagnose_source_includes_gettext_len_and_clipped() -> None:
    import inspect

    src = inspect.getsource(KbondSourceReader.diagnose)
    assert "gettext_len:" in src
    assert "clipped:" in src
