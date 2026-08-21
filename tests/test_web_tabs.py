from __future__ import annotations

import json
import sys
import time
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.license import sign_profile_dict
from app.profile import TraderProfile, apply_signed_profile, save_profile_raw
from app.runtime_status import RuntimeStatus, write_runtime_status
from app.web import server as web_server


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
        yield_prefix=3.0,
        mode=2,
        kbond_chat_title="room",
        sent_after="exit",
        message_template="{instrument} {confirm_token} ㅎㅈ",
        profile_version=2,
    )
    base.update(kwargs)
    return TraderProfile(**base)


def _asgi_json(app, method: str, path: str, *, headers=None, body=None):
    import asyncio

    payload = b"" if body is None else json.dumps(body).encode("utf-8")
    hdrs = list(headers or [])
    if body is not None and not any(k.lower() == b"content-type" for k, _ in hdrs):
        hdrs.append((b"content-type", b"application/json"))
    if not any(k.lower() == b"host" for k, _ in hdrs):
        hdrs.append((b"host", b"127.0.0.1:8765"))
    messages: list[dict] = []

    async def receive():
        return {"type": "http.request", "body": payload, "more_body": False}

    async def send(message):
        messages.append(message)

    asyncio.run(
        app(
            {
                "type": "http",
                "asgi": {"version": "3.0"},
                "http_version": "1.1",
                "method": method,
                "scheme": "http",
                "path": path,
                "raw_path": path.encode("ascii"),
                "query_string": b"",
                "headers": hdrs,
                "client": ("127.0.0.1", 123),
                "server": ("127.0.0.1", 80),
            },
            receive,
            send,
        )
    )
    start = next(m for m in messages if m["type"] == "http.response.start")
    raw = b"".join(m.get("body", b"") for m in messages if m["type"] == "http.response.body")
    data = json.loads(raw.decode("utf-8") or "{}")
    return int(start["status"]), data


@pytest.fixture()
def trader_env(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    monkeypatch.setenv("LOCALAPPDATA", str(tmp_path))
    monkeypatch.setenv("KBOND_DEPLOY_MODE", "pilot")
    monkeypatch.setenv("KBOND_ADMIN_URL", "https://example.invalid")
    import app.crypto_sign as crypto
    import app.paths as paths_mod

    monkeypatch.setattr(crypto, "data_dir", lambda: tmp_path)
    monkeypatch.setattr(paths_mod, "data_dir", lambda: tmp_path)
    monkeypatch.setattr(web_server, "LOCAL_TOKEN", "test-token")
    write_runtime_status(RuntimeStatus(state="STOPPED", engine_version="0.3.0"))
    return tmp_path


def test_profile_active_unauthorized(trader_env: Path):
    app = web_server.create_app()
    code, data = _asgi_json(
        app,
        "GET",
        "/api/profile/active",
        headers=[(b"x-kbond-token", b"test-token")],
    )
    assert code == 200
    assert data["authorized"] is False
    assert data["profile"] is None
    assert data["defaults"]["profile_name"] == ""


def test_profile_active_authorized(trader_env: Path):
    profile = _valid_profile()
    apply_signed_profile(profile, sign_profile_dict(profile))
    from app.license import (
        LicenseLease,
        activate_device,
        load_or_create_device,
        save_lease,
    )
    from app.machine import load_or_create_machine
    from app.crypto_sign import admin_sign_payload

    machine = load_or_create_machine()
    device = load_or_create_device(machine.machine_id)
    activate_device(device)
    lease = LicenseLease(
        device_id=device.device_id,
        trader_id="t1",
        profile_version=2,
        min_engine_version="0.1.0",
        expires_at=time.time() + 3600,
        enabled=True,
        machine_id=machine.machine_id,
    )
    lease.signature = admin_sign_payload(lease.to_payload())
    save_lease(lease)

    app = web_server.create_app()
    code, data = _asgi_json(
        app,
        "GET",
        "/api/profile/active",
        headers=[(b"x-kbond-token", b"test-token")],
    )
    assert code == 200
    assert data["authorized"] is True
    assert data["profile"]["profile_name"] == "t1"
    assert data["lease_ok"] is True

    code2, status = _asgi_json(
        app,
        "GET",
        "/api/status",
        headers=[(b"x-kbond-token", b"test-token")],
    )
    assert code2 == 200
    assert status["authorized"] is True
    assert status["lease_ok"] is True


def test_require_stopped_allows_sent(trader_env: Path):
    profile = _valid_profile()
    apply_signed_profile(profile, sign_profile_dict(profile))
    write_runtime_status(RuntimeStatus(state="SENT", engine_version="0.3.0"))
    from app.license import (
        LicenseLease,
        activate_device,
        load_or_create_device,
        save_lease,
    )
    from app.machine import load_or_create_machine
    from app.crypto_sign import admin_sign_payload

    machine = load_or_create_machine()
    device = load_or_create_device(machine.machine_id)
    activate_device(device)
    lease = LicenseLease(
        device_id=device.device_id,
        trader_id="t1",
        profile_version=2,
        min_engine_version="0.1.0",
        expires_at=time.time() + 3600,
        enabled=True,
        machine_id=machine.machine_id,
    )
    lease.signature = admin_sign_payload(lease.to_payload())
    save_lease(lease)

    app = web_server.create_app()
    body = profile.to_dict()
    body["runtime_only"] = True
    body["threshold"] = -50
    code, data = _asgi_json(
        app,
        "POST",
        "/api/profile",
        headers=[(b"x-kbond-token", b"test-token")],
        body=body,
    )
    assert code == 200
    assert data["mode"] == "runtime"


def test_test_click_requires_auth(trader_env: Path):
    app = web_server.create_app()
    code, data = _asgi_json(
        app,
        "POST",
        "/api/test-click",
        headers=[(b"x-kbond-token", b"test-token")],
        body={},
    )
    assert code == 400
    assert "not authorized" in str(data.get("detail", "")).lower()


def test_test_click_ok(trader_env: Path, monkeypatch: pytest.MonkeyPatch):
    profile = _valid_profile()
    apply_signed_profile(profile, sign_profile_dict(profile))
    calls = {"n": 0}

    def _fake_click(_cfg):
        calls["n"] += 1

    monkeypatch.setattr("send.click_only", _fake_click)
    app = web_server.create_app()
    code, data = _asgi_json(
        app,
        "POST",
        "/api/test-click",
        headers=[(b"x-kbond-token", b"test-token")],
        body={},
    )
    assert code == 200
    assert data["ok"] is True
    assert calls["n"] == 1
