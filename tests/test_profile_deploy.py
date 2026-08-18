from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.adapter import config_from_profile, slot_from_profile
from app.machine import MachineCalibration, save_machine, load_or_create_machine
from app.profile import (
    ProfileError,
    TraderProfile,
    apply_signed_profile,
    load_profile,
    save_profile,
    save_profile_draft,
)
from app.side_map import required_side_from_looking_for
from app.license import (
    activate_device,
    engine_meets_minimum,
    issue_lease,
    load_or_create_device,
    load_profile_signature,
    verify_lease_for_start,
    verify_signed_profile,
    LicenseError,
)
from app.runtime_status import RuntimeStatus, read_runtime_status, write_runtime_status
from app.audit import append_audit, iter_audit
from app.crypto_sign import admin_sign_payload
from app.deploy_mode import get_deploy_mode
from app.cred_protect import protect_secret, unprotect_secret


def _valid_profile(**kwargs) -> TraderProfile:
    base = dict(
        profile_name="t1",
        trader_id="t1",
        instrument="25-11",
        looking_for="BID",
        required_qty=100,
        threshold=-1000.0,
        excel_workbook=r"C:\Trading\bond.xlsm",
        excel_sheet="Sheet1",
        yield_input_cell="D19",
        pnl_cell="F22",
        yield_prefix=3.5,
        mode=2,
        kbond_chat_title="room",
        sent_after="exit",
        message_template="{instrument} {confirm_token} ㅎㅈ",
    )
    base.update(kwargs)
    return TraderProfile(**base)


def test_side_map() -> None:
    assert required_side_from_looking_for("BID") == "BUY"
    assert required_side_from_looking_for("OFFER") == "SELL"


def test_slot_from_profile() -> None:
    profile = _valid_profile()
    slot = slot_from_profile(profile, 3.5)
    assert slot.instrument == "25-11"
    assert slot.looking_for == "BID"
    assert slot.required_side == "BUY"
    assert slot.qty_abs == 100
    assert slot.input_cell == "D19"
    assert slot.pnl_cell == "F22"


def test_profile_rejects_bare_workbook_name() -> None:
    profile = _valid_profile(excel_workbook="bond.xlsm")
    with pytest.raises(ProfileError):
        profile.validate()


def test_profile_atomic_save(tmp_path: Path) -> None:
    path = tmp_path / "profile.json"
    p = _valid_profile(profile_version=3)
    saved = save_profile(p, path)
    assert saved.profile_version == 4
    loaded = load_profile(path)
    assert loaded.instrument == "25-11"
    assert loaded.profile_version == 4


def test_config_from_profile_mode2() -> None:
    profile = _valid_profile(mode=2)
    machine = MachineCalibration(machine_id="m1", send_input_x=0.2, send_input_y=0.9)
    cfg = config_from_profile(profile, machine)
    assert cfg.mode == 2
    assert cfg.send_input_x == 0.2
    assert cfg.excel_workbook.endswith("bond.xlsm")
    assert cfg.excel_status_cell == ""


def test_runtime_status_roundtrip(tmp_path: Path) -> None:
    path = tmp_path / "runtime_status.json"
    write_runtime_status(
        RuntimeStatus(state="WATCHING", watcher_pid=1, instrument="25-10"),
        path=path,
    )
    got = read_runtime_status(path)
    assert got.state == "WATCHING"
    assert got.watcher_pid == 1
    assert got.heartbeat_at


def test_audit_append(tmp_path: Path) -> None:
    path = tmp_path / "audit.jsonl"
    eid = append_audit("PROFILE_SAVED", {"trader_id": "t"}, path=path)
    rows = list(iter_audit(path))
    assert rows[0]["event_id"] == eid
    assert rows[0]["event"] == "PROFILE_SAVED"


def test_lease_gate(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "dev")
    import app.license as lic
    import app.crypto_sign as crypto

    monkeypatch.setattr(lic, "device_path", lambda: tmp_path / "device.json")
    monkeypatch.setattr(lic, "lease_path", lambda: tmp_path / "lease.json")
    monkeypatch.setattr(crypto, "data_dir", lambda: tmp_path)

    machine = MachineCalibration(machine_id="m1")
    device = load_or_create_device(machine.machine_id)
    assert device.credential_blob
    assert "credential" not in device.to_dict() or not device.to_dict().get("credential")
    profile = _valid_profile()
    with pytest.raises(LicenseError):
        verify_lease_for_start(device=device, profile=profile, lease=None)
    activate_device(device)
    lease = issue_lease(device=device, profile=profile, ttl_seconds=60)
    assert "engine_version" not in lease.to_payload()
    assert lease.min_engine_version
    verify_lease_for_start(device=device, profile=profile, lease=lease)
    assert engine_meets_minimum("0.2.0", "0.1.0")
    assert not engine_meets_minimum("0.1.0", "0.2.0")


def test_pilot_forbids_local_issue_lease(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "pilot")
    import app.license as lic
    import app.crypto_sign as crypto

    monkeypatch.setattr(lic, "device_path", lambda: tmp_path / "device.json")
    monkeypatch.setattr(lic, "lease_path", lambda: tmp_path / "lease.json")
    monkeypatch.setattr(crypto, "data_dir", lambda: tmp_path)

    device = load_or_create_device("m1")
    activate_device(device)
    with pytest.raises(LicenseError, match="pilot"):
        issue_lease(device=device, profile=_valid_profile(), ttl_seconds=60)


def test_signed_profile_required_in_pilot(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "pilot")
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    import app.crypto_sign as crypto

    monkeypatch.setattr(crypto, "data_dir", lambda: tmp_path)
    profile = _valid_profile()
    with pytest.raises(LicenseError):
        verify_signed_profile(profile, None)
    sig = admin_sign_payload(profile.to_dict())
    verify_signed_profile(profile, sig)
    bad = dict(profile.to_dict())
    bad["threshold"] = -99999
    from app.profile import TraderProfile as TP

    with pytest.raises(LicenseError):
        verify_signed_profile(TP.from_dict(bad), sig)


def test_deploy_mode_env_dev(monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.delenv("KBOND_DEPLOY_MODE", raising=False)
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "dev")
    # Cannot easily un-freeze; just assert env path when not frozen
    from app import deploy_mode

    monkeypatch.setattr(deploy_mode, "is_frozen_binary", lambda: False)
    monkeypatch.setattr(deploy_mode, "DEPLOY_MODE_BUILD", None)
    assert get_deploy_mode() == "dev"


def test_deploy_mode_build_pilot_wins(monkeypatch: pytest.MonkeyPatch) -> None:
    from app import deploy_mode

    monkeypatch.setattr(deploy_mode, "is_frozen_binary", lambda: False)
    monkeypatch.setattr(deploy_mode, "DEPLOY_MODE_BUILD", "pilot")
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "dev")
    assert get_deploy_mode() == "pilot"


def test_dpapi_roundtrip() -> None:
    blob, method = protect_secret(b"secret-bytes-32!!!!!!!!!!!!!!!!")
    assert method == "dpapi"
    assert unprotect_secret(blob, method) == b"secret-bytes-32!!!!!!!!!!!!!!!!"


def test_policy_check_disable(
    tmp_path: Path, monkeypatch: pytest.MonkeyPatch
) -> None:
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "dev")
    monkeypatch.delenv("KBOND_ADMIN_URL", raising=False)
    import app.license as lic
    import app.crypto_sign as crypto
    import app.policy_poll as poll

    monkeypatch.setattr(lic, "device_path", lambda: tmp_path / "device.json")
    monkeypatch.setattr(lic, "lease_path", lambda: tmp_path / "lease.json")
    monkeypatch.setattr(crypto, "data_dir", lambda: tmp_path)

    from app.paths import profile_path
    from app.profile import save_profile_raw

    profile = _valid_profile(profile_version=1)
    save_profile_raw(profile, profile_path())
    # unsigned ok in dev
    device = load_or_create_device("m1")
    activate_device(device)
    issue_lease(device=device, profile=profile, ttl_seconds=60)
    assert poll.check_policy_or_stop() is None
    device.disabled = True
    lic.save_device(device)
    assert poll.check_policy_or_stop() is not None


def test_apply_signed_profile(tmp_path: Path, monkeypatch: pytest.MonkeyPatch) -> None:
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "dev")
    import app.crypto_sign as crypto
    import app.paths as paths_mod

    monkeypatch.setattr(crypto, "data_dir", lambda: tmp_path)

    profile = _valid_profile(profile_version=2)
    sig = admin_sign_payload(profile.to_dict())
    apply_signed_profile(profile, sig)
    assert load_profile().profile_version == 2
    assert load_profile_signature()
