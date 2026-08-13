from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from enum import Enum
from hashlib import sha1
from typing import Optional


class AppStatus(str, Enum):
    IDLE = "IDLE"
    STARTING = "STARTING"
    WATCHING = "WATCHING"
    QUOTE_FOUND = "QUOTE_FOUND"
    CALCULATING = "CALCULATING"
    NO_TRIGGER = "NO_TRIGGER"
    TRIGGERED = "TRIGGERED"
    SENDING = "SENDING"
    SENT = "SENT"
    DONE = "DONE"
    STOPPED = "STOPPED"
    ERROR = "ERROR"


@dataclass(frozen=True)
class Quote:
    instrument: str
    raw_line: str
    raw_token: str
    yield_value: float
    side: str
    timestamp: Optional[datetime] = None
    sender: Optional[str] = None

    @property
    def fingerprint(self) -> str:
        return sha1(self.raw_line.encode("utf-8")).hexdigest()


@dataclass(frozen=True)
class TriggerResult:
    triggered: bool
    reason: str
    pnl: float
    quote: Optional[Quote] = None


@dataclass
class WatcherSession:
    processed_fingerprints: set[str] = field(default_factory=set)
    status: AppStatus = AppStatus.IDLE
