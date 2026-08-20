from __future__ import annotations

import json
import uuid
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Any, Optional

from .paths import machine_path


class MachineError(ValueError):
    pass


@dataclass
class MachineCalibration:
    machine_id: str = ""
    send_input_x: float = 0.5
    send_input_y: float = 0.9

    def validate(self) -> None:
        if not (self.machine_id or "").strip():
            raise MachineError("machine_id is required")
        for key in ("send_input_x", "send_input_y"):
            value = float(getattr(self, key))
            if not 0.0 <= value <= 1.0:
                raise MachineError(f"{key} must be between 0 and 1")

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "MachineCalibration":
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


def load_or_create_machine(path: Optional[Path] = None) -> MachineCalibration:
    target = path or machine_path()
    if target.is_file():
        data = json.loads(target.read_text(encoding="utf-8"))
        machine = MachineCalibration.from_dict(data if isinstance(data, dict) else {})
        machine.validate()
        return machine
    machine = MachineCalibration(machine_id=str(uuid.uuid4()))
    machine.validate()
    _atomic_write_json(target, machine.to_dict())
    return machine


def save_machine(machine: MachineCalibration, path: Optional[Path] = None) -> None:
    machine.validate()
    _atomic_write_json(path or machine_path(), machine.to_dict())
