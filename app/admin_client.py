from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import time
from typing import Any, Optional
from urllib import error, request

from app import ENGINE_VERSION
from app.deploy_mode import is_pilot
from app.license import (
    DeviceRecord,
    LicenseLease,
    get_device_credential_bytes,
    save_lease,
)
from app.profile import TraderProfile


class AdminClientError(RuntimeError):
    pass


def _require_admin_url(admin_url: str) -> str:
    url = (admin_url or "").strip().rstrip("/")
    if not url:
        raise AdminClientError("KBOND_ADMIN_URL is empty")
    if is_pilot() and not url.lower().startswith("https://"):
        raise AdminClientError("pilot mode requires HTTPS KBOND_ADMIN_URL")
    return url


def _auth_headers(device: DeviceRecord) -> dict[str, str]:
    ts = str(int(time.time()))
    secret = get_device_credential_bytes(device)
    mac = hmac.new(secret, f"{device.device_id}:{ts}".encode("utf-8"), hashlib.sha256).hexdigest()
    return {
        "Content-Type": "application/json",
        "X-Device-Id": device.device_id,
        "X-Machine-Id": device.machine_id,
        "X-Device-Timestamp": ts,
        "X-Device-Auth": mac,
        "X-Credential-Protection": device.credential_protection,
        "X-Engine-Version": ENGINE_VERSION,
    }


def _post_json(url: str, payload: dict[str, Any], device: DeviceRecord, timeout: float = 5.0) -> dict[str, Any]:
    body = json.dumps(payload).encode("utf-8")
    req = request.Request(url, data=body, headers=_auth_headers(device), method="POST")
    try:
        with request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise AdminClientError(f"admin HTTP {exc.code}: {detail}") from exc
    except (error.URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise AdminClientError(f"admin request failed: {exc}") from exc


def _get_json(url: str, device: DeviceRecord, timeout: float = 5.0) -> dict[str, Any]:
    req = request.Request(url, headers=_auth_headers(device), method="GET")
    try:
        with request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode("utf-8"))
    except error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise AdminClientError(f"admin HTTP {exc.code}: {detail}") from exc
    except (error.URLError, TimeoutError, json.JSONDecodeError) as exc:
        raise AdminClientError(f"admin request failed: {exc}") from exc


def refresh_lease_from_admin(
    admin_url: str,
    device: DeviceRecord,
    profile: TraderProfile,
    timeout: float = 5.0,
) -> LicenseLease:
    url = _require_admin_url(admin_url) + "/api/lease"
    data = _post_json(
        url,
        {
            "device_id": device.device_id,
            "machine_id": device.machine_id,
            "trader_id": profile.trader_id or profile.profile_name,
            "profile_version": profile.profile_version,
            "credential_protection": device.credential_protection,
        },
        device,
        timeout=timeout,
    )
    lease = LicenseLease.from_dict(data)
    save_lease(lease)
    return lease


def fetch_policy(
    admin_url: str,
    device: DeviceRecord,
    trader_id: str,
    timeout: float = 5.0,
) -> dict[str, Any]:
    url = _require_admin_url(admin_url) + f"/api/devices/{device.device_id}/policy"
    return _get_json(url, device, timeout=timeout)


def submit_profile_draft(
    admin_url: str,
    device: DeviceRecord,
    profile: TraderProfile,
    timeout: float = 5.0,
) -> dict[str, Any]:
    url = _require_admin_url(admin_url) + "/api/profile/submit"
    return _post_json(
        url,
        {"profile": profile.to_dict()},
        device,
        timeout=timeout,
    )


def fetch_signed_profile(
    admin_url: str,
    device: DeviceRecord,
    trader_id: str,
    timeout: float = 5.0,
) -> dict[str, Any]:
    url = _require_admin_url(admin_url) + f"/api/profile/current?trader_id={trader_id}"
    return _get_json(url, device, timeout=timeout)


def upload_audit_batch(
    admin_url: str,
    device: DeviceRecord,
    rows: list[dict[str, Any]],
    timeout: float = 5.0,
) -> dict[str, Any]:
    url = _require_admin_url(admin_url) + "/api/audit/ingest"
    return _post_json(url, {"events": rows}, device, timeout=timeout)


def register_device(
    admin_url: str,
    device: DeviceRecord,
    trader_id: str,
    timeout: float = 5.0,
) -> dict[str, Any]:
    """Send public device id + protection method; Admin stores auth verifier separately on activate."""
    url = _require_admin_url(admin_url) + "/api/devices/register"
    # Admin DB stores secret hex (Admin-only) to verify device HMAC auth.
    secret = get_device_credential_bytes(device)
    return _post_json(
        url,
        {
            "device_id": device.device_id,
            "machine_id": device.machine_id,
            "trader_id": trader_id,
            "credential_secret_hex": secret.hex(),
            "credential_hash": hashlib.sha256(secret).hexdigest(),
            "credential_protection": device.credential_protection,
            "engine_version": ENGINE_VERSION,
        },
        device,
        timeout=timeout,
    )
