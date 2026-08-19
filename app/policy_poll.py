from __future__ import annotations

import logging
import os
import threading
from typing import Any, Optional

from app.admin_client import AdminClientError, fetch_policy, fetch_signed_profile, refresh_lease_from_admin
from app.audit import append_audit
from app.defaults import DEFAULTS
from app.demo_expiry import check_demo_expiry
from app.license import (
    LicenseError,
    load_lease,
    load_or_create_device,
    load_profile_signature,
    verify_policy_snapshot,
    verify_signed_profile,
)
from app.machine import load_or_create_machine
from app.profile import (
    ProfileError,
    TraderProfile,
    apply_signed_profile,
    load_profile,
    load_profile_draft,
)
from app.runtime_status import read_runtime_status

logger = logging.getLogger("kbond_watcher")

_poll_stop = threading.Event()
_poll_thread: Optional[threading.Thread] = None
_last_profile_sync: dict[str, Any] = {}


def get_last_profile_sync() -> dict[str, Any]:
    return dict(_last_profile_sync)


def check_policy_or_stop() -> Optional[str]:
    """
    Returns error message if policy requires stop; None if OK.
    Never called from Quote→Excel→Send path.
    """
    try:
        check_demo_expiry()
        profile = load_profile()
        machine = load_or_create_machine()
        device = load_or_create_device(machine.machine_id)
        lease = load_lease()
        admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()

        enabled = None
        min_engine = None
        if admin_url:
            try:
                policy = fetch_policy(admin_url, device, profile.trader_id or profile.profile_name)
                enabled = policy.get("enabled")
                min_engine = policy.get("min_engine_version")
                if policy.get("disabled") is True:
                    enabled = False
                try:
                    lease = refresh_lease_from_admin(admin_url, device, profile)
                except AdminClientError as exc:
                    logger.info("lease refresh skipped: %s", exc)
            except AdminClientError as exc:
                logger.info("policy fetch failed (using cache): %s", exc)

        verify_signed_profile(profile, load_profile_signature())
        verify_policy_snapshot(
            device=device,
            profile=profile,
            lease=lease,
            enabled=False if enabled is False else None,
            min_engine_version=min_engine,
        )
        return None
    except (LicenseError, AdminClientError, Exception) as exc:
        return str(exc)


def try_pull_approved_profile() -> None:
    """
    When STOPPED: if Admin has a newer signed profile, apply locally.
    One lightweight GET per policy_poll_seconds — negligible load.
    """
    global _last_profile_sync
    admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
    if not admin_url:
        return
    state = read_runtime_status().state
    if state not in {"STOPPED", "ERROR", "", "DONE", "SENT"}:
        return
    try:
        try:
            draft = load_profile_draft()
            trader_id = draft.trader_id or draft.profile_name
        except ProfileError:
            try:
                active = load_profile()
                trader_id = active.trader_id or active.profile_name
            except ProfileError:
                return
        machine = load_or_create_machine()
        device = load_or_create_device(machine.machine_id)
        data = fetch_signed_profile(admin_url, device, trader_id)
        remote = TraderProfile.from_dict(data["profile"])
        remote_sig = str(data["signature"])
        local_ver = 0
        local_sig = load_profile_signature()
        try:
            local_ver = int(load_profile().profile_version)
        except ProfileError:
            local_ver = 0
        if local_sig and int(remote.profile_version) <= local_ver:
            _last_profile_sync = {
                "ok": True,
                "pending": False,
                "profile_version": local_ver,
                "message": "signed profile up to date",
            }
            return
        apply_signed_profile(remote, remote_sig)
        try:
            refresh_lease_from_admin(admin_url, device, remote)
        except AdminClientError as exc:
            logger.info("lease refresh after profile pull skipped: %s", exc)
        _last_profile_sync = {
            "ok": True,
            "pending": False,
            "applied": True,
            "profile_version": remote.profile_version,
            "message": f"auto-applied signed profile_version={remote.profile_version}",
        }
        logger.info("auto-applied signed profile version=%s", remote.profile_version)
    except AdminClientError as exc:
        msg = str(exc)
        pending = "404" in msg or "not found" in msg.lower()
        _last_profile_sync = {
            "ok": True if pending else False,
            "pending": pending,
            "message": "waiting for Admin approval" if pending else msg,
        }
    except Exception as exc:
        _last_profile_sync = {"ok": False, "pending": False, "message": str(exc)}
        logger.info("profile pull skipped: %s", exc)


def _poll_loop() -> None:
    from app.controller import stop_watcher

    while not _poll_stop.wait(DEFAULTS.policy_poll_seconds):
        try:
            try_pull_approved_profile()
        except Exception as exc:
            logger.info("profile pull error: %s", exc)
        status = read_runtime_status().state
        if status not in {"WATCHING", "EXCEL_WAIT", "STARTING", "SENT"}:
            continue
        reason = check_policy_or_stop()
        if reason:
            logger.warning("policy poll revoke: %s", reason)
            append_audit(
                "LICENSE_REJECTED",
                {"error_message": reason, "source": "policy_poll"},
            )
            try:
                stop_watcher()
            except Exception as exc:
                logger.error("policy soft stop failed: %s", exc)


def start_policy_poller() -> None:
    global _poll_thread
    if _poll_thread and _poll_thread.is_alive():
        return
    _poll_stop.clear()

    def _run() -> None:
        try:
            try_pull_approved_profile()
        except Exception as exc:
            logger.info("initial profile pull: %s", exc)
        _poll_loop()

    _poll_thread = threading.Thread(target=_run, name="policy-poller", daemon=True)
    _poll_thread.start()


def stop_policy_poller() -> None:
    _poll_stop.set()
