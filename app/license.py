from __future__ import annotations

import base64
import json
import os
import time
import uuid
from dataclasses import asdict, dataclass, fields
from pathlib import Path
from typing import Any, Optional

from . import ENGINE_VERSION
from .cred_protect import ProtectionMethod, protect_secret, unprotect_secret
from .crypto_sign import admin_sign_payload, verify_admin_signature
from .deploy_mode import is_dev, is_pilot
from .paths import device_path, lease_path, profile_path
from .profile import TraderProfile


class LicenseError(RuntimeError):
    pass


@dataclass
class DeviceRecord:
    machine_id: str
    device_id: str
    activated: bool = False
    credential_blob: str = ""
    credential_protection: ProtectionMethod = "dpapi"
    disabled: bool = False

    def to_dict(self) -> dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "DeviceRecord":
        known = {f.name for f in fields(cls)}
        raw = {k: v for k, v in data.items() if k in known}
        if "credential_protection" in raw and raw["credential_protection"] not in {
            "dpapi",
            "tpm",
            "plaintext_dev",
        }:
            raw["credential_protection"] = "dpapi"
        return cls(**raw)


@dataclass
class LicenseLease:
    device_id: str
    trader_id: str
    profile_version: int
    min_engine_version: str
    expires_at: float
    enabled: bool = True
    machine_id: str = ""
    signature: str = ""

    def to_payload(self) -> dict[str, Any]:
        return {
            "device_id": self.device_id,
            "trader_id": self.trader_id,
            "profile_version": int(self.profile_version),
            "min_engine_version": self.min_engine_version,
            "expires_at": float(self.expires_at),
            "enabled": bool(self.enabled),
        }

    def to_dict(self) -> dict[str, Any]:
        data = self.to_payload()
        data["machine_id"] = self.machine_id
        data["signature"] = self.signature
        return data

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "LicenseLease":
        enabled = data.get("enabled")
        if enabled is None:
            enabled = not bool(data.get("disabled", False))
        return cls(
            device_id=str(data.get("device_id") or ""),
            trader_id=str(data.get("trader_id") or ""),
            profile_version=int(data.get("profile_version") or 0),
            min_engine_version=str(data.get("min_engine_version") or "0.0.0"),
            expires_at=float(data.get("expires_at") or 0),
            enabled=bool(enabled),
            machine_id=str(data.get("machine_id") or ""),
            signature=str(data.get("signature") or ""),
        )


def _version_tuple(text: str) -> tuple[int, ...]:
    parts = []
    for bit in (text or "0").split("."):
        try:
            parts.append(int(bit))
        except ValueError:
            parts.append(0)
    return tuple(parts) or (0,)


def engine_meets_minimum(current: str, minimum: str) -> bool:
    return _version_tuple(current) >= _version_tuple(minimum)


def _atomic_write(path: Path, payload: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    tmp.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(path)


def load_or_create_device(machine_id: str) -> DeviceRecord:
    path = device_path()
    if path.is_file():
        data = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(data, dict):
            data = {}
        # migrate legacy plaintext
        if data.get("credential") and not data.get("credential_blob"):
            secret = str(data["credential"]).encode("utf-8")
            blob, method = protect_secret(secret)
            data["credential_blob"] = blob
            data["credential_protection"] = method
            data.pop("credential", None)
            _atomic_write(path, data)
        return DeviceRecord.from_dict(data)

    secret = os.urandom(32)
    blob, method = protect_secret(secret)
    record = DeviceRecord(
        machine_id=machine_id,
        device_id=str(uuid.uuid4()),
        activated=False,
        credential_blob=blob,
        credential_protection=method,
    )
    _atomic_write(path, record.to_dict())
    return record


def save_device(record: DeviceRecord) -> None:
    payload = record.to_dict()
    payload.pop("credential", None)
    _atomic_write(device_path(), payload)


def get_device_credential_bytes(record: DeviceRecord) -> bytes:
    if not record.credential_blob:
        raise LicenseError("device credential missing")
    try:
        return unprotect_secret(record.credential_blob, record.credential_protection)
    except Exception as exc:
        raise LicenseError(f"failed to unprotect device credential: {exc}") from exc


def activate_device(record: DeviceRecord) -> DeviceRecord:
    if is_pilot() and not record.activated:
        # Pilot: activation must come from Admin; local flag alone is not enough
        # but we still allow setting activated after Admin acknowledges.
        pass
    record.activated = True
    save_device(record)
    return record


def issue_lease(
    *,
    device: DeviceRecord,
    profile: TraderProfile,
    ttl_seconds: float = 3600.0,
    min_engine_version: str = "0.1.0",
    enabled: bool = True,
) -> LicenseLease:
    """Local lease issuance — DEV ONLY. Pilot must use Admin-signed leases."""
    if is_pilot():
        raise LicenseError("local lease issuance forbidden in pilot mode")
    lease = LicenseLease(
        device_id=device.device_id,
        trader_id=profile.trader_id or profile.profile_name,
        profile_version=int(profile.profile_version),
        min_engine_version=min_engine_version,
        expires_at=time.time() + float(ttl_seconds),
        enabled=enabled,
        machine_id=device.machine_id,
    )
    lease.signature = admin_sign_payload(lease.to_payload())
    _atomic_write(lease_path(), lease.to_dict())
    return lease


def save_lease(lease: LicenseLease) -> None:
    _atomic_write(lease_path(), lease.to_dict())


def load_lease() -> Optional[LicenseLease]:
    path = lease_path()
    if not path.is_file():
        return None
    try:
        data = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(data, dict):
        return None
    return LicenseLease.from_dict(data)


def profile_signature_path(path: Optional[Path] = None) -> Path:
    base = path or profile_path()
    return base.with_suffix(base.suffix + ".sig")


def load_profile_signature(path: Optional[Path] = None) -> Optional[str]:
    sig_path = profile_signature_path(path)
    if not sig_path.is_file():
        return None
    return sig_path.read_text(encoding="ascii").strip()


def save_profile_signature(signature: str, path: Optional[Path] = None) -> None:
    sig_path = profile_signature_path(path)
    sig_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = sig_path.with_suffix(sig_path.suffix + ".tmp")
    tmp.write_text(signature.strip() + "\n", encoding="ascii")
    tmp.replace(sig_path)


def verify_lease_for_start(
    *,
    device: DeviceRecord,
    profile: TraderProfile,
    lease: Optional[LicenseLease],
) -> None:
    if not device.activated:
        raise LicenseError("device is not activated")
    if device.disabled:
        raise LicenseError("device is remotely disabled")
    if lease is None:
        raise LicenseError("license lease missing")
    if not verify_admin_signature(lease.to_payload(), lease.signature):
        raise LicenseError("license lease signature invalid")
    if lease.device_id != device.device_id:
        raise LicenseError("license lease device mismatch")
    if not lease.enabled:
        raise LicenseError("license lease disabled")
    if int(lease.profile_version) != int(profile.profile_version):
        raise LicenseError(
            f"lease profile_version {lease.profile_version} != "
            f"local {profile.profile_version}"
        )
    if not engine_meets_minimum(ENGINE_VERSION, lease.min_engine_version):
        raise LicenseError(
            f"engine {ENGINE_VERSION} below minimum {lease.min_engine_version}"
        )
    if time.time() > float(lease.expires_at):
        raise LicenseError("license lease expired")


def verify_policy_snapshot(
    *,
    device: DeviceRecord,
    profile: TraderProfile,
    lease: Optional[LicenseLease],
    enabled: Optional[bool] = None,
    min_engine_version: Optional[str] = None,
) -> None:
    """Used by Controller poll — same fail-closed rules."""
    if enabled is False or device.disabled:
        raise LicenseError("remotely disabled")
    verify_lease_for_start(device=device, profile=profile, lease=lease)
    if min_engine_version and not engine_meets_minimum(ENGINE_VERSION, min_engine_version):
        raise LicenseError(
            f"engine {ENGINE_VERSION} below minimum {min_engine_version}"
        )


def sign_profile_dict(profile: dict[str, Any]) -> str:
    return admin_sign_payload(profile)


def verify_signed_profile(profile: TraderProfile, signature: Optional[str]) -> None:
    if is_dev() and not signature:
        return
    if not signature:
        raise LicenseError("profile signature required")
    if not verify_admin_signature(profile.to_dict(), signature):
        raise LicenseError("profile signature invalid")


# Back-compat aliases used by older imports
def sign_payload(payload: dict[str, Any]) -> str:
    return admin_sign_payload(payload)


def verify_signature(payload: dict[str, Any], signature: str) -> bool:
    return verify_admin_signature(payload, signature)
