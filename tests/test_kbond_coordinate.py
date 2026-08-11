"""Unit tests for K-Bond relative client coordinates."""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from kbond_controller import compute_client_click  # noqa: E402


def test_relative_coordinate_example() -> None:
    # client rect 1000 x 800, WIN_X=0.5, WIN_Y=0.4 → (500, 320)
    x, y = compute_client_click(1000, 800, 0.5, 0.4)
    assert (x, y) == (500, 320)


@pytest.mark.parametrize(
    "w,h,wx,wy,expected",
    [
        (1000, 800, 0.0, 0.0, (0, 0)),
        (1000, 800, 1.0, 1.0, (1000, 800)),
        (640, 480, 0.25, 0.5, (160, 240)),
    ],
)
def test_relative_coordinates(
    w: int, h: int, wx: float, wy: float, expected: tuple[int, int]
) -> None:
    assert compute_client_click(w, h, wx, wy) == expected


def test_invalid_ratios() -> None:
    with pytest.raises(ValueError):
        compute_client_click(100, 100, -0.1, 0.5)
    with pytest.raises(ValueError):
        compute_client_click(100, 100, 0.5, 1.1)


def test_select_best_hwnd_prefers_title_then_preferred_pid() -> None:
    from kbond_controller import select_best_hwnd

    candidates = [
        (1, "Other", 10_000, 100),
        (2, "KBond Login", 500, 200),
        (3, "KBond Main", 9_000, 100),
    ]
    # Title filter narrows to KBond*, then preferred_pid=100 picks Main over Login area
    assert select_best_hwnd(candidates, title_contains="KBond", preferred_pid=100) == 3
    # Without preferred pid, largest titled wins
    assert select_best_hwnd(candidates, title_contains="KBond") == 3
    # Exact title-only smaller window still beats non-matching large when title set
    assert select_best_hwnd(
        [(1, "Other", 99_999, 1), (2, "KBond Login", 100, 2)],
        title_contains="Login",
    ) == 2


def test_select_best_hwnd_empty() -> None:
    from kbond_controller import select_best_hwnd

    assert select_best_hwnd([]) is None


def test_process_name_tokens_include_axis_family() -> None:
    from kbond_controller import _process_name_tokens

    tokens = _process_name_tokens("axis.exe")
    assert "axis" in tokens
    assert "kbond" in tokens

