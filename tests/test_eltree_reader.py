from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from source.eltree import TreeCandidate, normalize_lines, pick_chat_tree


def test_normalize_lines_strips_and_dedupes() -> None:
    assert normalize_lines(["  a\n", "a", "b\nb", "", "  "]) == ["a", "b"]


def test_normalize_lines_preserves_order() -> None:
    assert normalize_lines(["c", "a", "b", "a"]) == ["c", "a", "b"]


def test_pick_chat_tree_prefers_right_large() -> None:
    parent_left, parent_right = 0, 1000
    left_room = TreeCandidate(hwnd=1, left=10, top=10, right=200, bottom=800)
    right_small = TreeCandidate(hwnd=2, left=600, top=10, right=900, bottom=200)
    right_large = TreeCandidate(hwnd=3, left=550, top=10, right=980, bottom=900)
    chosen = pick_chat_tree(
        [left_room, right_small, right_large],
        parent_left,
        parent_right,
    )
    assert chosen is not None
    assert chosen.hwnd == 3


def test_pick_chat_tree_rejects_left_only() -> None:
    parent_left, parent_right = 0, 1000
    left_room = TreeCandidate(hwnd=1, left=10, top=10, right=400, bottom=800)
    chosen = pick_chat_tree([left_room], parent_left, parent_right)
    assert chosen is None


def test_center_x_ratio() -> None:
    cand = TreeCandidate(hwnd=1, left=700, top=0, right=900, bottom=100)
    assert abs(cand.center_x_ratio(0, 1000) - 0.8) < 1e-9
