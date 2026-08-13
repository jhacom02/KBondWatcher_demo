from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from send import relative_point  # noqa: E402


def test_relative_point_center() -> None:
    assert relative_point(0, 0, 1000, 800, 0.5, 0.5) == (500, 400)


def test_relative_point_input() -> None:
    assert relative_point(100, 200, 400, 600, 0.5, 0.5) == (300, 500)
