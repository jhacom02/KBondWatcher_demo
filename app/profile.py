from __future__ import annotations

import json
import re
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Any, Optional

from . import PROFILE_SCHEMA_VERSION
from .paths import profile_draft_path, profile_path
from .side_map import normalize_looking_for

_CELL_RE = re.compile(r"^[A-Za-z]{1,3}\d{1,7}$")
ALLOWED_INSTRUMENTS = frozenset({"25-10", "25-4", "25-8", "25-5", "25-11"})
THRESHOLD_OPS = frozenset({"<=", ">="})

POLICY_KEYS = (
    "profile_name",
    "trader_id",
    "profile_version",
    "profile_schema_version",
    "kbond_chat_title",
    "mode",
    "sent_after",
    "excel_workbook",
    "excel_sheet",
    "message_template",
)

RUNTIME_KEYS = (
    "instrument",
    "looking_for",
    "required_qty",
    "threshold",
    "threshold_op",
    "yield_prefix",
    "yield_input_cell",
    "pnl_cell",
)


class ProfileError(ValueError):
    pass


def policy_payload(profile: TraderProfile | dict[str, Any]) -> dict[str, Any]:
    data = profile.to_dict() if isinstance(profile, TraderProfile) else dict(profile)
    out: dict[str, Any] = {}
    for key in POLICY_KEYS:
        if key == "profile_version":
            out[key] = int(data.get(key) or 0)
        elif key == "profile_schema_version":
            out[key] = int(data.get(key) or PROFILE_SCHEMA_VERSION)
        elif key == "mode":
            out[key] = int(data.get(key) or 0)
        elif key == "sent_after":
            out[key] = str(data.get(key) or "").strip().lower()
        else:
            out[key] = data.get(key) if data.get(key) is not None else ""
            if key in {
                "profile_name",
                "trader_id",
                "kbond_chat_title",
                "excel_workbook",
                "excel_sheet",
                "message_template",
            }:
                out[key] = str(out[key]).strip() if key != "excel_workbook" else str(out[key])
    return out


def policy_fields_equal(a: TraderProfile | dict[str, Any], b: TraderProfile | dict[str, Any]) -> bool:
    return policy_payload(a) == policy_payload(b)


def prefs_snapshot(profile: TraderProfile | dict[str, Any]) -> dict[str, Any]:
    data = profile.to_dict() if isinstance(profile, TraderProfile) else dict(profile)
    snap = policy_payload(data)
    for key in RUNTIME_KEYS:
        snap[key] = data.get(key)
    return snap


@dataclass
class TraderProfile:
    profile_name: str = ""
    profile_version: int = 1
    profile_schema_version: int = PROFILE_SCHEMA_VERSION
    trader_id: str = ""
    instrument: str = ""
    looking_for: str = "BID"
    required_qty: int = 100
    threshold: float = 0.0
    threshold_op: str = "<="
    excel_workbook: str = ""
    excel_sheet: str = ""
    yield_input_cell: str = "A1"
    pnl_cell: str = "B1"
    yield_prefix: float = 3.0
    yield_prefix_cell: str = ""
    mode: int = 2
    kbond_chat_title: str = "[채권] 블커본드"
    sent_after: str = "exit"
    message_template: str = "{instrument} {confirm_token} ㅎㅈ"

    def validate(self) -> None:
        if self.profile_schema_version != PROFILE_SCHEMA_VERSION:
            raise ProfileError(
                f"unsupported profile_schema_version "
                f"{self.profile_schema_version} (need {PROFILE_SCHEMA_VERSION})"
            )
        if not (self.profile_name or "").strip():
            raise ProfileError("profile_name is required")
        inst = (self.instrument or "").strip()
        if inst not in ALLOWED_INSTRUMENTS:
            raise ProfileError(f"instrument must be one of {sorted(ALLOWED_INSTRUMENTS)}")
        normalize_looking_for(self.looking_for)
        if int(self.required_qty) <= 0:
            raise ProfileError("required_qty must be > 0")
        try:
            float(self.threshold)
        except (TypeError, ValueError) as exc:
            raise ProfileError("threshold must be numeric") from exc
        op = (self.threshold_op or "").strip()
        if op not in THRESHOLD_OPS:
            raise ProfileError("threshold_op must be <= or >=")
        if not (self.excel_workbook or "").strip():
            raise ProfileError("excel_workbook FullName is required")
        if "\\" not in self.excel_workbook and "/" not in self.excel_workbook:
            raise ProfileError(
                "excel_workbook must be a FullName path, not a bare Name"
            )
        if not (self.excel_sheet or "").strip():
            raise ProfileError("excel_sheet is required")
        for key in ("yield_input_cell", "pnl_cell"):
            cell = getattr(self, key)
            if not _CELL_RE.match((cell or "").strip()):
                raise ProfileError(f"{key} must be a cell like D19, got {cell!r}")
        mode = int(self.mode)
        if mode not in (1, 2, 3):
            raise ProfileError("mode must be 1, 2, or 3")
        if mode in (1, 2) and not (self.kbond_chat_title or "").strip():
            raise ProfileError("kbond_chat_title is required for MODE 1/2")
        sent = (self.sent_after or "").strip().lower()
        if sent not in {"exit", "loop"}:
            raise ProfileError("sent_after must be exit or loop")
        if mode == 1 and sent != "exit":
            raise ProfileError("sent_after must be exit when mode is 1")
        if not (self.message_template or "").strip():
            raise ProfileError("message_template is required")

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "TraderProfile":
        known = {f.name for f in fields(cls)}
        kwargs = {k: v for k, v in data.items() if k in known}
        return cls(**kwargs)


def _atomic_write_json(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )
    tmp.replace(path)


def load_profile(path: Optional[Path] = None) -> TraderProfile:
    target = path or profile_path()
    if not target.is_file():
        raise ProfileError(f"profile not found: {target}")
    data = json.loads(target.read_text(encoding="utf-8"))
    if not isinstance(data, dict):
        raise ProfileError("profile must be a JSON object")
    profile = TraderProfile.from_dict(data)
    profile.validate()
    return profile


def save_profile(profile: TraderProfile, path: Optional[Path] = None) -> TraderProfile:
    profile.validate()
    target = path or profile_path()
    bump = TraderProfile(**{**profile.to_dict(), "profile_version": int(profile.profile_version) + 1})
    bump.validate()
    _atomic_write_json(target, bump.to_dict())
    return bump


def save_profile_raw(profile: TraderProfile, path: Optional[Path] = None) -> None:
    profile.validate()
    _atomic_write_json(path or profile_path(), profile.to_dict())


def save_profile_draft(profile: TraderProfile) -> TraderProfile:
    profile.validate()
    draft = TraderProfile(
        **{**profile.to_dict(), "profile_version": int(profile.profile_version)}
    )
    _atomic_write_json(profile_draft_path(), draft.to_dict())
    return draft


def load_profile_draft() -> TraderProfile:
    return load_profile(profile_draft_path())


def apply_signed_profile(profile: TraderProfile, signature: str) -> None:
    from .license import save_profile_signature, verify_signed_profile

    profile.validate()
    verify_signed_profile(profile, signature)
    save_profile_raw(profile)
    save_profile_signature(signature)
