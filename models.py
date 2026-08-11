"""Shared dataclasses and status enums for the FORESTBOND watcher."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from typing import Optional


class AppStatus(str, Enum):
    IDLE = "IDLE"
    STARTING = "STARTING"
    WATCHING = "WATCHING"
    QUOTE_FOUND = "QUOTE_FOUND"
    CALCULATING = "CALCULATING"
    NO_TRIGGER = "NO_TRIGGER"
    TRIGGERED = "TRIGGERED"
    PREFILLING = "PREFILLING"
    READY_TO_SUBMIT = "READY_TO_SUBMIT"
    DONE = "DONE"
    STOPPED = "STOPPED"
    ERROR = "ERROR"


@dataclass(frozen=True)
class Quote:
    instrument: str
    raw_line: str
    raw_token: str
    yield_value: float
    side: str  # BUY | SELL
    timestamp: Optional[datetime] = None
    sender: Optional[str] = None

    @property
    def fingerprint(self) -> str:
        from hashlib import sha1

        return sha1(self.raw_line.encode("utf-8")).hexdigest()


@dataclass(frozen=True)
class TriggerResult:
    triggered: bool
    reason: str
    pnl: float
    quote: Optional[Quote] = None


@dataclass
class KBondWindowInfo:
    pid: int
    process_name: str
    hwnd: int
    title: str
    window_rect: tuple[int, int, int, int]
    client_rect: tuple[int, int, int, int]
    click_client: tuple[int, int]
    click_screen: tuple[int, int]


@dataclass
class WatcherSession:
    """In-memory state for a single one-shot watcher run."""

    processed_fingerprints: set[str] = field(default_factory=set)
    status: AppStatus = AppStatus.IDLE
