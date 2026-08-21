from __future__ import annotations

import json
import uuid
from datetime import datetime
from pathlib import Path
from typing import Any, Optional

from .paths import audit_path, audit_upload_status_path

AUDIT_EVENTS = frozenset(
    {
        "PROFILE_SAVED",
        "PROFILE_RUNTIME_SAVED",
        "WATCHER_STARTED",
        "WATCHER_STOPPED",
        "TRIGGER_RESULT",
        "SEND_ATTEMPT",
        "SEND_RESULT",
        "ERROR",
        "LICENSE_REJECTED",
        "PROFILE_REJECTED",
    }
)


def append_audit(event: str, fields: Optional[dict[str, Any]] = None, path: Optional[Path] = None) -> str:
    if event not in AUDIT_EVENTS:
        raise ValueError(f"unsupported audit event: {event}")
    event_id = str(uuid.uuid4())
    record: dict[str, Any] = {
        "event_id": event_id,
        "timestamp": datetime.now().astimezone().isoformat(timespec="seconds"),
        "event": event,
    }
    if fields:
        record.update(fields)
    target = path or audit_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    with target.open("a", encoding="utf-8") as fh:
        fh.write(json.dumps(record, ensure_ascii=False) + "\n")
    return event_id


def iter_audit(path: Optional[Path] = None, *, after_id: Optional[str] = None):
    target = path or audit_path()
    if not target.is_file():
        return
    seen_after = after_id is None
    with target.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            try:
                row = json.loads(line)
            except json.JSONDecodeError:
                continue
            if not seen_after:
                if row.get("event_id") == after_id:
                    seen_after = True
                continue
            yield row


def write_audit_upload_status(payload: dict[str, Any]) -> None:
    path = audit_upload_status_path()
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(path)


def read_audit_upload_status() -> dict[str, Any]:
    path = audit_upload_status_path()
    if not path.is_file():
        return {"ok": True, "pending": 0, "stale": False}
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {"ok": False, "pending": 0, "stale": True, "error": "status unreadable"}
    if not isinstance(data, dict):
        return {"ok": False, "stale": True}
    return data
