from __future__ import annotations

import json
from dataclasses import asdict, dataclass, fields
from datetime import datetime, timezone
from pathlib import Path
from typing import Any, Optional

from .paths import runtime_status_path


def _now_iso() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


@dataclass
class RuntimeStatus:
    state: str = "STOPPED"
    watcher_pid: Optional[int] = None
    heartbeat_at: Optional[str] = None
    instrument: str = ""
    looking_for: str = ""
    last_quote: Optional[str] = None
    last_pnl: Optional[float] = None
    threshold: Optional[float] = None
    last_action: Optional[str] = None
    last_error: Optional[str] = None
    profile_version: Optional[int] = None
    engine_version: str = ""

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "RuntimeStatus":
        known = {f.name for f in fields(cls)}
        kwargs = {k: v for k, v in data.items() if k in known}
        return cls(**kwargs)


def write_runtime_status(status: RuntimeStatus, path: Optional[Path] = None) -> None:
    target = path or runtime_status_path()
    status.heartbeat_at = _now_iso()
    target.parent.mkdir(parents=True, exist_ok=True)
    tmp = target.with_suffix(target.suffix + ".tmp")
    tmp.write_text(
        json.dumps(status.to_dict(), ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    tmp.replace(target)


def read_runtime_status(path: Optional[Path] = None) -> RuntimeStatus:
    target = path or runtime_status_path()
    if not target.is_file():
        return RuntimeStatus(state="STOPPED")
    try:
        data = json.loads(target.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return RuntimeStatus(state="STOPPED", last_error="corrupt runtime_status.json")
    if not isinstance(data, dict):
        return RuntimeStatus(state="STOPPED")
    return RuntimeStatus.from_dict(data)
