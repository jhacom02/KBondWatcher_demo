from __future__ import annotations

import logging
import os
import threading
from typing import Optional

from app.admin_client import AdminClientError, fetch_policy, refresh_lease_from_admin
from app.audit import append_audit
from app.defaults import DEFAULTS
from app.license import (
    LicenseError,
    load_lease,
    load_or_create_device,
    load_profile_signature,
    verify_policy_snapshot,
    verify_signed_profile,
)
from app.machine import load_or_create_machine
from app.profile import load_profile

logger = logging.getLogger("kbond_watcher")

_poll_stop = threading.Event()
_poll_thread: Optional[threading.Thread] = None


def check_policy_or_stop() -> Optional[str]:
    """
    Returns error message if policy requires stop; None if OK.
    Never called from Quote→Excel→Send path.
    """
    try:
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


def _poll_loop() -> None:
    from app.controller import stop_watcher
    from app.runtime_status import read_runtime_status

    while not _poll_stop.wait(DEFAULTS.policy_poll_seconds):
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
    _poll_thread = threading.Thread(target=_poll_loop, name="policy-poller", daemon=True)
    _poll_thread.start()


def stop_policy_poller() -> None:
    _poll_stop.set()
