from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from message_sender import relative_point  # noqa: E402


def test_relative_point_chat_tab() -> None:
    assert relative_point(0, 0, 1000, 800, 0.08, 0.15) == (80, 120)


def test_relative_point_input() -> None:
    assert relative_point(100, 200, 400, 600, 0.5, 0.9) == (300, 740)
