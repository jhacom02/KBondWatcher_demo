from __future__ import annotations

from core.trigger import LOOKING_BID, LOOKING_OFFER


def required_side_from_looking_for(looking_for: str) -> str:
    looking = (looking_for or "").strip().upper()
    if looking == LOOKING_BID:
        return "BUY"
    if looking == LOOKING_OFFER:
        return "SELL"
    raise ValueError(f"looking_for must be BID or OFFER, got {looking_for!r}")


def normalize_looking_for(looking_for: str) -> str:
    looking = (looking_for or "").strip().upper()
    if looking not in {LOOKING_BID, LOOKING_OFFER}:
        raise ValueError(f"looking_for must be BID or OFFER, got {looking_for!r}")
    return looking
