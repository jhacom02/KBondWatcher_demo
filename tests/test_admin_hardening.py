from __future__ import annotations

import asyncio
import hashlib
import hmac
import json
import sys
import time
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from app.license import sign_profile_dict
from admin import server as admin_server
from admin.server import fmt_pilot_expires, fmt_ts, init_db


@pytest.fixture()
def admin_db(tmp_path: Path, monkeypatch: pytest.MonkeyPatch):
    db = tmp_path / "admin.db"
    monkeypatch.setattr(admin_server, "DB_PATH", db)
    init_db()
    return db


def _auth_headers(device_id: str, secret: bytes) -> list[tuple[bytes, bytes]]:
    ts = str(int(time.time()))
    mac = hmac.new(secret, f"{device_id}:{ts}".encode("utf-8"), hashlib.sha256).hexdigest()
    return [
        (b"x-device-id", device_id.encode("utf-8")),
        (b"x-device-timestamp", ts.encode("utf-8")),
        (b"x-device-auth", mac.encode("utf-8")),
        (b"content-type", b"application/json"),
    ]


def _seed_device(secret: bytes, *, trader_id: str = "t1", enabled: int = 1) -> None:
    with admin_server._connect() as conn:
        conn.execute(
            "INSERT INTO machines(device_id, machine_id, trader_id, last_seen, engine_version, "
            "credential_hash, credential_protection, enabled, activated, pilot_expires_at) "
            "VALUES(?,?,?,?,?,?,?,?,?,?)",
            (
                "dev1",
                "m1",
                trader_id,
                time.time(),
                "0.2.0",
                secret.hex(),
                "dpapi",
                enabled,
                1,
                time.time() + 7 * 24 * 3600,
            ),
        )


def _seed_signed_trader(trader_id: str = "t1", *, disabled: int = 0) -> dict:
    profile = {
        "profile_name": trader_id,
        "trader_id": trader_id,
        "profile_version": 2,
        "profile_schema_version": 1,
        "instrument": "25-10",
        "looking_for": "BID",
        "required_qty": 100,
        "threshold": -1000,
        "threshold_op": "<=",
        "excel_workbook": r"C:\Trading\bond.xlsm",
        "excel_sheet": "Sheet1",
        "yield_input_cell": "D19",
        "pnl_cell": "F22",
        "yield_prefix": 3,
        "mode": 2,
        "kbond_chat_title": "room",
        "sent_after": "exit",
        "message_template": "{instrument} {confirm_token} ㅎㅈ",
    }
    sig = sign_profile_dict(profile)
    with admin_server._connect() as conn:
        conn.execute(
            "INSERT INTO traders(trader_id, disabled, min_engine_version, profile_json, "
            "profile_signature, profile_version, updated_at) VALUES(?,?,?,?,?,?,?)",
            (
                trader_id,
                disabled,
                "0.1.0",
                json.dumps(profile, ensure_ascii=False),
                sig,
                2,
                time.time(),
            ),
        )
    return profile


def _asgi_json(
    app,
    method: str,
    path: str,
    *,
    headers: list[tuple[bytes, bytes]] | None = None,
    body: dict | None = None,
    query: str = "",
) -> tuple[int, dict]:
    payload = b"" if body is None else json.dumps(body).encode("utf-8")
    hdrs = list(headers or [])
    if body is not None and not any(k.lower() == b"content-type" for k, _ in hdrs):
        hdrs.append((b"content-type", b"application/json"))
    scope = {
        "type": "http",
        "asgi": {"version": "3.0"},
        "http_version": "1.1",
        "method": method,
        "scheme": "http",
        "path": path,
        "raw_path": path.encode("ascii"),
        "query_string": query.encode("ascii"),
        "headers": hdrs,
        "client": ("127.0.0.1", 123),
        "server": ("test", 80),
    }
    messages: list[dict] = []

    async def receive():
        return {"type": "http.request", "body": payload, "more_body": False}

    async def send(message):
        messages.append(message)

    asyncio.run(app(scope, receive, send))
    start = next(m for m in messages if m["type"] == "http.response.start")
    body_parts = [m.get("body", b"") for m in messages if m["type"] == "http.response.body"]
    raw = b"".join(body_parts)
    data = json.loads(raw.decode("utf-8") or "{}")
    return int(start["status"]), data


def test_fmt_ts_and_pilot() -> None:
    assert fmt_ts(None) == "—"
    assert fmt_ts("") == "—"
    ts = 1_700_000_000.0
    assert len(fmt_ts(ts)) == 16
    assert "expired" in fmt_pilot_expires(ts, now=ts + 10)
    assert "d left" in fmt_pilot_expires(ts + 3 * 86400, now=ts)


def test_revoke_clears_signature(admin_db: Path) -> None:
    secret = b"x" * 32
    _seed_device(secret)
    _seed_signed_trader()
    app = admin_server.create_app()
    status, data = _asgi_json(app, "POST", "/api/traders/t1/revoke-profile")
    assert status == 200
    assert data.get("ok") is True
    with admin_server._connect() as conn:
        row = conn.execute(
            "SELECT profile_json, profile_signature, profile_version FROM traders WHERE trader_id=?",
            ("t1",),
        ).fetchone()
    assert row["profile_json"] is None
    assert row["profile_signature"] is None
    assert int(row["profile_version"]) == 2
    status, _ = _asgi_json(
        app,
        "GET",
        "/api/profile/current",
        headers=_auth_headers("dev1", secret),
        query="trader_id=t1",
    )
    assert status == 404


def test_lease_requires_signed_profile(admin_db: Path) -> None:
    secret = b"y" * 32
    _seed_device(secret)
    with admin_server._connect() as conn:
        conn.execute(
            "INSERT INTO traders(trader_id, disabled, min_engine_version, profile_version, updated_at) "
            "VALUES(?,?,?,?,?)",
            ("t1", 0, "0.1.0", 1, time.time()),
        )
    app = admin_server.create_app()
    status, data = _asgi_json(
        app,
        "POST",
        "/api/lease",
        headers=_auth_headers("dev1", secret),
        body={"trader_id": "t1", "profile_version": 1, "machine_id": "m1"},
    )
    assert status == 403
    assert "signed profile" in str(data.get("detail"))


def test_lease_refuses_disabled_trader(admin_db: Path) -> None:
    secret = b"z" * 32
    _seed_device(secret)
    _seed_signed_trader(disabled=1)
    app = admin_server.create_app()
    status, data = _asgi_json(
        app,
        "POST",
        "/api/lease",
        headers=_auth_headers("dev1", secret),
        body={"trader_id": "t1", "profile_version": 2, "machine_id": "m1"},
    )
    assert status == 403
    assert "disabled" in str(data.get("detail"))


def test_lease_ok_with_signed_profile(admin_db: Path) -> None:
    secret = b"w" * 32
    _seed_device(secret)
    _seed_signed_trader()
    app = admin_server.create_app()
    status, data = _asgi_json(
        app,
        "POST",
        "/api/lease",
        headers=_auth_headers("dev1", secret),
        body={"trader_id": "t1", "profile_version": 2, "machine_id": "m1"},
    )
    assert status == 200
    assert data["enabled"] is True
    assert data["profile_version"] == 2
    assert data["signature"]
