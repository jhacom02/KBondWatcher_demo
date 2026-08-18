from __future__ import annotations

import os
from pathlib import Path


def data_dir() -> Path:
    base = os.environ.get("LOCALAPPDATA") or str(Path.home() / "AppData" / "Local")
    path = Path(base) / "KBondWatcher"
    path.mkdir(parents=True, exist_ok=True)
    return path


def profile_path() -> Path:
    return data_dir() / "profile.json"


def profile_draft_path() -> Path:
    return data_dir() / "profile.draft.json"


def machine_path() -> Path:
    return data_dir() / "machine.json"


def runtime_status_path() -> Path:
    return data_dir() / "runtime_status.json"


def audit_path() -> Path:
    return data_dir() / "audit.jsonl"


def lease_path() -> Path:
    return data_dir() / "lease.json"


def device_path() -> Path:
    return data_dir() / "device.json"


def local_token_path() -> Path:
    return data_dir() / "web_token.txt"


def audit_cursor_path() -> Path:
    return data_dir() / "audit_cursor.txt"


def audit_upload_status_path() -> Path:
    return data_dir() / "audit_upload_status.json"


def stop_flag_path() -> Path:
    return data_dir() / "watcher.stop"


def logs_dir() -> Path:
    path = data_dir() / "logs"
    path.mkdir(parents=True, exist_ok=True)
    return path
