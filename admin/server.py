from __future__ import annotations

import hashlib
import hmac
import json
import os
import sqlite3
import time
from pathlib import Path
from typing import Any, Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse
from fastapi.templating import Jinja2Templates

from app import ENGINE_VERSION
from app.crypto_sign import admin_sign_payload, export_public_key_b64

DB_PATH = Path(__file__).resolve().parent / "admin.db"
TEMPLATES = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))

_AUTH_SKEW_SECONDS = 300
_DEFAULT_LEASE_WINDOW = 7 * 24 * 3600


def lease_window_seconds() -> int:
    raw = (os.environ.get("KBOND_LEASE_TTL_SECONDS") or "").strip()
    if not raw:
        return _DEFAULT_LEASE_WINDOW
    try:
        value = int(raw)
    except ValueError:
        return _DEFAULT_LEASE_WINDOW
    return value if value > 0 else _DEFAULT_LEASE_WINDOW


def compute_lease_expires_at(now: float, pilot_expires_at: float, ttl: Optional[int] = None) -> float:
    window = float(ttl if ttl is not None else lease_window_seconds())
    return min(now + window, float(pilot_expires_at))


def _connect() -> sqlite3.Connection:
    DB_PATH.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    return conn


def _migrate(conn: sqlite3.Connection) -> None:
    cols = {r["name"] for r in conn.execute("PRAGMA table_info(machines)").fetchall()}
    alters = [
        ("credential_hash", "TEXT"),
        ("credential_protection", "TEXT DEFAULT 'dpapi'"),
        ("enabled", "INTEGER NOT NULL DEFAULT 1"),
        ("activated", "INTEGER NOT NULL DEFAULT 0"),
        ("last_lease_at", "REAL"),
        ("last_audit_at", "REAL"),
        ("pilot_expires_at", "REAL"),
    ]
    for name, decl in alters:
        if name not in cols:
            conn.execute(f"ALTER TABLE machines ADD COLUMN {name} {decl}")

    tcols = {r["name"] for r in conn.execute("PRAGMA table_info(traders)").fetchall()}
    if "profile_signature" not in tcols:
        conn.execute("ALTER TABLE traders ADD COLUMN profile_signature TEXT")
    if "draft_json" not in tcols:
        conn.execute("ALTER TABLE traders ADD COLUMN draft_json TEXT")


def init_db() -> None:
    with _connect() as conn:
        conn.executescript(
            """
            CREATE TABLE IF NOT EXISTS traders (
              trader_id TEXT PRIMARY KEY,
              disabled INTEGER NOT NULL DEFAULT 0,
              min_engine_version TEXT NOT NULL DEFAULT '0.1.0',
              profile_json TEXT,
              profile_signature TEXT,
              draft_json TEXT,
              profile_version INTEGER NOT NULL DEFAULT 0,
              updated_at REAL
            );
            CREATE TABLE IF NOT EXISTS machines (
              device_id TEXT PRIMARY KEY,
              machine_id TEXT,
              trader_id TEXT,
              last_seen REAL,
              engine_version TEXT,
              credential_hash TEXT,
              credential_protection TEXT DEFAULT 'dpapi',
              enabled INTEGER NOT NULL DEFAULT 1,
              activated INTEGER NOT NULL DEFAULT 0,
              last_lease_at REAL,
              last_audit_at REAL,
              pilot_expires_at REAL
            );
            CREATE TABLE IF NOT EXISTS audit_events (
              event_id TEXT PRIMARY KEY,
              timestamp TEXT,
              event TEXT,
              trader_id TEXT,
              machine_id TEXT,
              payload TEXT
            );
            """
        )
        _migrate(conn)


def _ensure_pilot_expires_at(
    conn: sqlite3.Connection, device_id: str, existing: Optional[sqlite3.Row]
) -> float:
    """Set pilot_expires_at once on first register/lease; never extend."""
    now = time.time()
    if existing is not None and existing["pilot_expires_at"]:
        return float(existing["pilot_expires_at"])
    expires = now + lease_window_seconds()
    conn.execute(
        "UPDATE machines SET pilot_expires_at=? WHERE device_id=? AND "
        "(pilot_expires_at IS NULL OR pilot_expires_at=0)",
        (expires, device_id),
    )
    return expires


def _touch_seen(
    conn: sqlite3.Connection,
    *,
    device_id: str,
    machine_id: str = "",
    trader_id: str = "",
    engine_version: str = "",
) -> None:
    conn.execute(
        "INSERT INTO machines(device_id, machine_id, trader_id, last_seen, engine_version) "
        "VALUES(?,?,?,?,?) "
        "ON CONFLICT(device_id) DO UPDATE SET last_seen=excluded.last_seen, "
        "machine_id=COALESCE(NULLIF(excluded.machine_id,''), machines.machine_id), "
        "trader_id=COALESCE(NULLIF(excluded.trader_id,''), machines.trader_id), "
        "engine_version=COALESCE(NULLIF(excluded.engine_version,''), machines.engine_version)",
        (device_id, machine_id, trader_id, time.time(), engine_version),
    )


def _verify_device_auth(request: Request, conn: sqlite3.Connection) -> sqlite3.Row:
    device_id = request.headers.get("X-Device-Id") or ""
    ts = request.headers.get("X-Device-Timestamp") or ""
    mac = request.headers.get("X-Device-Auth") or ""
    if not device_id or not ts or not mac:
        raise HTTPException(401, "device auth headers required")
    try:
        ts_i = int(ts)
    except ValueError as exc:
        raise HTTPException(401, "bad timestamp") from exc
    if abs(time.time() - ts_i) > _AUTH_SKEW_SECONDS:
        raise HTTPException(401, "auth timestamp skew")
    row = conn.execute("SELECT * FROM machines WHERE device_id=?", (device_id,)).fetchone()
    if row is None or not row["credential_hash"]:
        raise HTTPException(401, "device not registered")
    # credential_hash column stores Admin-only secret hex for HMAC verify
    try:
        secret = bytes.fromhex(str(row["credential_hash"]))
    except ValueError as exc:
        raise HTTPException(401, "device credential corrupt") from exc
    expected = hmac.new(secret, f"{device_id}:{ts}".encode("utf-8"), hashlib.sha256).hexdigest()
    if not hmac.compare_digest(expected, mac):
        raise HTTPException(401, "device auth failed")
    if int(row["enabled"] or 0) == 0:
        raise HTTPException(403, "device disabled")
    if int(row["activated"] or 0) == 0:
        raise HTTPException(403, "device not activated")
    return row


def create_app() -> FastAPI:
    init_db()
    app = FastAPI(title="KBondWatcher Admin")

    @app.get("/", response_class=HTMLResponse)
    async def home(request: Request):
        with _connect() as conn:
            traders = conn.execute("SELECT * FROM traders ORDER BY trader_id").fetchall()
            machines = conn.execute("SELECT * FROM machines ORDER BY last_seen DESC").fetchall()
            audits = conn.execute(
                "SELECT * FROM audit_events ORDER BY timestamp DESC LIMIT 100"
            ).fetchall()
        return TEMPLATES.TemplateResponse(
            request,
            "index.html",
            {
                "traders": traders,
                "machines": machines,
                "audits": audits,
                "engine_version": ENGINE_VERSION,
                "public_key": export_public_key_b64(),
            },
        )

    @app.post("/api/devices/register")
    async def api_register(request: Request):
        data = await request.json()
        device_id = str(data.get("device_id") or "")
        machine_id = str(data.get("machine_id") or "")
        trader_id = str(data.get("trader_id") or "")
        secret_material = str(data.get("credential_secret_hex") or "")
        if not secret_material:
            raise HTTPException(400, "credential_secret_hex required")
        protection = str(data.get("credential_protection") or "dpapi")
        engine_version = str(data.get("engine_version") or ENGINE_VERSION)
        if not device_id or not secret_material:
            raise HTTPException(400, "device_id and credential_secret_hex required")
        with _connect() as conn:
            existing = conn.execute(
                "SELECT * FROM machines WHERE device_id=?", (device_id,)
            ).fetchone()
            if existing and existing["credential_hash"]:
                # Re-registration must prove possession of existing secret
                _verify_device_auth(request, conn)
            activated = 1
            if existing and existing["activated"] is not None:
                activated = int(existing["activated"])
            if existing is None:
                activated = 1
            pilot_expires = None
            if existing is not None and existing["pilot_expires_at"]:
                pilot_expires = float(existing["pilot_expires_at"])
            elif existing is None:
                pilot_expires = time.time() + lease_window_seconds()
            conn.execute(
                "INSERT INTO machines(device_id, machine_id, trader_id, last_seen, engine_version, "
                "credential_hash, credential_protection, enabled, activated, pilot_expires_at) "
                "VALUES(?,?,?,?,?,?,?,?,?,?) "
                "ON CONFLICT(device_id) DO UPDATE SET "
                "machine_id=excluded.machine_id, trader_id=excluded.trader_id, "
                "last_seen=excluded.last_seen, engine_version=excluded.engine_version, "
                "credential_hash=excluded.credential_hash, "
                "credential_protection=excluded.credential_protection, "
                "pilot_expires_at=COALESCE(machines.pilot_expires_at, excluded.pilot_expires_at)",
                (
                    device_id,
                    machine_id,
                    trader_id,
                    time.time(),
                    engine_version,
                    secret_material,
                    protection,
                    1,
                    activated,
                    pilot_expires,
                ),
            )
            # Backfill absolute window for legacy rows missing pilot_expires_at
            row_full = conn.execute(
                "SELECT * FROM machines WHERE device_id=?", (device_id,)
            ).fetchone()
            pilot_expires_at = _ensure_pilot_expires_at(conn, device_id, row_full)
            if trader_id:
                conn.execute(
                    "INSERT INTO traders(trader_id, updated_at) VALUES(?,?) "
                    "ON CONFLICT(trader_id) DO UPDATE SET updated_at=excluded.updated_at",
                    (trader_id, time.time()),
                )
            row = conn.execute(
                "SELECT activated, enabled, credential_protection, pilot_expires_at "
                "FROM machines WHERE device_id=?",
                (device_id,),
            ).fetchone()
        return {
            "ok": True,
            "activated": bool(row["activated"]),
            "enabled": bool(row["enabled"]),
            "credential_protection": row["credential_protection"],
            "pilot_expires_at": row["pilot_expires_at"] or pilot_expires_at,
        }

    @app.get("/api/devices/{device_id}/policy")
    async def api_policy(device_id: str, request: Request):
        with _connect() as conn:
            device = _verify_device_auth(request, conn)
            trader_id = device["trader_id"] or ""
            trader = None
            if trader_id:
                trader = conn.execute(
                    "SELECT * FROM traders WHERE trader_id=?", (trader_id,)
                ).fetchone()
            _touch_seen(
                conn,
                device_id=device_id,
                machine_id=request.headers.get("X-Machine-Id") or "",
                engine_version=request.headers.get("X-Engine-Version") or "",
            )
        enabled = bool(device["enabled"]) and not bool(trader["disabled"] if trader else 0)
        return {
            "device_id": device_id,
            "enabled": enabled,
            "disabled": not enabled,
            "min_engine_version": (trader["min_engine_version"] if trader else "0.1.0"),
            "credential_protection": device["credential_protection"] or "dpapi",
            "activated": bool(device["activated"]),
        }

    @app.post("/api/lease")
    async def api_lease(request: Request):
        data = await request.json()
        with _connect() as conn:
            device = _verify_device_auth(request, conn)
            trader_id = str(data.get("trader_id") or device["trader_id"] or "")
            profile_version = int(data.get("profile_version") or 0)
            row = conn.execute(
                "SELECT * FROM traders WHERE trader_id=?", (trader_id,)
            ).fetchone()
            if row is None:
                conn.execute(
                    "INSERT INTO traders(trader_id, disabled, min_engine_version, profile_version, updated_at) "
                    "VALUES(?,?,?,?,?)",
                    (trader_id, 0, "0.1.0", profile_version or 1, time.time()),
                )
                row = conn.execute(
                    "SELECT * FROM traders WHERE trader_id=?", (trader_id,)
                ).fetchone()
            disabled = bool(row["disabled"]) or not bool(device["enabled"])
            min_engine = row["min_engine_version"] or "0.1.0"
            if profile_version <= 0:
                profile_version = int(row["profile_version"] or 1)
            if row["profile_version"] and int(row["profile_version"]) > 0:
                # Prefer Admin signed profile version when present
                if row["profile_json"]:
                    profile_version = int(row["profile_version"])
            now = time.time()
            pilot_expires_at = _ensure_pilot_expires_at(conn, device["device_id"], device)
            if now > float(pilot_expires_at):
                raise HTTPException(403, "pilot window expired; lease reissue refused")
            expires_at = compute_lease_expires_at(now, pilot_expires_at)
            payload = {
                "device_id": device["device_id"],
                "trader_id": trader_id,
                "profile_version": profile_version,
                "min_engine_version": min_engine,
                "expires_at": expires_at,
                "enabled": not disabled,
            }
            payload["signature"] = admin_sign_payload(
                {k: v for k, v in payload.items() if k != "signature"}
            )
            payload["machine_id"] = device["machine_id"] or data.get("machine_id") or ""
            payload["pilot_expires_at"] = pilot_expires_at
            conn.execute(
                "UPDATE machines SET last_lease_at=?, last_seen=?, trader_id=?, "
                "credential_protection=COALESCE(?, credential_protection) WHERE device_id=?",
                (
                    time.time(),
                    time.time(),
                    trader_id,
                    data.get("credential_protection"),
                    device["device_id"],
                ),
            )
        return payload

    @app.post("/api/profile/submit")
    async def api_profile_submit(request: Request):
        data = await request.json()
        profile = data.get("profile") or {}
        if not isinstance(profile, dict):
            raise HTTPException(400, "profile required")
        trader_id = str(profile.get("trader_id") or profile.get("profile_name") or "")
        if not trader_id:
            raise HTTPException(400, "trader_id required")
        with _connect() as conn:
            _verify_device_auth(request, conn)
            conn.execute(
                "INSERT INTO traders(trader_id, draft_json, updated_at) VALUES(?,?,?) "
                "ON CONFLICT(trader_id) DO UPDATE SET draft_json=excluded.draft_json, "
                "updated_at=excluded.updated_at",
                (trader_id, json.dumps(profile, ensure_ascii=False), time.time()),
            )
        return {"ok": True, "trader_id": trader_id, "status": "draft"}

    @app.post("/api/profile/approve/{trader_id}")
    async def api_profile_approve(trader_id: str):
        with _connect() as conn:
            row = conn.execute(
                "SELECT draft_json, profile_version FROM traders WHERE trader_id=?",
                (trader_id,),
            ).fetchone()
            if not row or not row["draft_json"]:
                raise HTTPException(404, "draft not found")
            profile = json.loads(row["draft_json"])
            version = int(row["profile_version"] or 0) + 1
            profile["trader_id"] = trader_id
            profile["profile_version"] = version
            signature = admin_sign_payload(profile)
            conn.execute(
                "UPDATE traders SET profile_json=?, profile_signature=?, profile_version=?, "
                "updated_at=?, draft_json=NULL WHERE trader_id=?",
                (
                    json.dumps(profile, ensure_ascii=False),
                    signature,
                    version,
                    time.time(),
                    trader_id,
                ),
            )
        return {"ok": True, "profile_version": version, "signature": signature, "profile": profile}

    @app.get("/api/profile/current")
    async def api_profile_current(request: Request, trader_id: str = ""):
        with _connect() as conn:
            device = _verify_device_auth(request, conn)
            tid = trader_id or device["trader_id"] or ""
            row = conn.execute(
                "SELECT profile_json, profile_signature, profile_version FROM traders WHERE trader_id=?",
                (tid,),
            ).fetchone()
        if not row or not row["profile_json"] or not row["profile_signature"]:
            raise HTTPException(404, "signed profile not found")
        return {
            "profile": json.loads(row["profile_json"]),
            "signature": row["profile_signature"],
            "profile_version": row["profile_version"],
        }

    @app.post("/api/audit/ingest")
    async def api_audit_ingest(request: Request):
        data = await request.json()
        events = data.get("events") or []
        accepted = 0
        duplicates = 0
        with _connect() as conn:
            try:
                device = _verify_device_auth(request, conn)
                device_id = device["device_id"]
            except HTTPException:
                # Allow unauthenticated ingest only if no device headers (legacy) — reject in pilot hardening
                raise
            for ev in events:
                event_id = ev.get("event_id")
                if not event_id:
                    continue
                try:
                    conn.execute(
                        "INSERT INTO audit_events(event_id, timestamp, event, trader_id, machine_id, payload) "
                        "VALUES(?,?,?,?,?,?)",
                        (
                            event_id,
                            ev.get("timestamp"),
                            ev.get("event"),
                            ev.get("trader_id"),
                            ev.get("machine_id"),
                            json.dumps(ev, ensure_ascii=False),
                        ),
                    )
                    accepted += 1
                except sqlite3.IntegrityError:
                    duplicates += 1
                    continue
            conn.execute(
                "UPDATE machines SET last_audit_at=?, last_seen=? WHERE device_id=?",
                (time.time(), time.time(), device_id),
            )
        return {"accepted": accepted, "duplicates": duplicates}

    @app.post("/api/traders/{trader_id}/disable")
    async def disable_trader(trader_id: str, request: Request):
        body = await request.json()
        disabled = 1 if body.get("disabled") else 0
        with _connect() as conn:
            conn.execute(
                "INSERT INTO traders(trader_id, disabled, updated_at) VALUES(?,?,?) "
                "ON CONFLICT(trader_id) DO UPDATE SET disabled=excluded.disabled, updated_at=excluded.updated_at",
                (trader_id, disabled, time.time()),
            )
        return {"ok": True, "disabled": bool(disabled)}

    @app.post("/api/devices/{device_id}/disable")
    async def disable_device(device_id: str, request: Request):
        body = await request.json()
        enabled = 0 if body.get("disabled") else 1
        if "enabled" in body:
            enabled = 1 if body.get("enabled") else 0
        with _connect() as conn:
            conn.execute(
                "UPDATE machines SET enabled=? WHERE device_id=?",
                (enabled, device_id),
            )
        return {"ok": True, "enabled": bool(enabled)}

    @app.post("/api/traders/{trader_id}/min-engine")
    async def min_engine(trader_id: str, request: Request):
        body = await request.json()
        version = str(body.get("min_engine_version") or "0.1.0")
        with _connect() as conn:
            conn.execute(
                "INSERT INTO traders(trader_id, min_engine_version, updated_at) VALUES(?,?,?) "
                "ON CONFLICT(trader_id) DO UPDATE SET min_engine_version=excluded.min_engine_version, "
                "updated_at=excluded.updated_at",
                (trader_id, version, time.time()),
            )
        return {"ok": True, "min_engine_version": version}

    @app.get("/api/traders/{trader_id}/profile")
    async def get_profile(trader_id: str):
        with _connect() as conn:
            row = conn.execute(
                "SELECT profile_json, profile_signature, profile_version FROM traders WHERE trader_id=?",
                (trader_id,),
            ).fetchone()
        if not row or not row["profile_json"]:
            raise HTTPException(404, "profile not found")
        return {
            "profile": json.loads(row["profile_json"]),
            "signature": row["profile_signature"],
            "profile_version": row["profile_version"],
        }

    @app.put("/api/traders/{trader_id}/profile")
    async def put_profile(trader_id: str, request: Request):
        data = await request.json()
        data["trader_id"] = trader_id
        data["profile_version"] = int(data.get("profile_version") or 0) + 1
        signature = admin_sign_payload(data)
        with _connect() as conn:
            conn.execute(
                "INSERT INTO traders(trader_id, profile_json, profile_signature, profile_version, updated_at) "
                "VALUES(?,?,?,?,?) "
                "ON CONFLICT(trader_id) DO UPDATE SET profile_json=excluded.profile_json, "
                "profile_signature=excluded.profile_signature, "
                "profile_version=excluded.profile_version, updated_at=excluded.updated_at",
                (
                    trader_id,
                    json.dumps(data, ensure_ascii=False),
                    signature,
                    data["profile_version"],
                    time.time(),
                ),
            )
        return {"ok": True, "profile_version": data["profile_version"], "signature": signature}

    return app


def run_admin(host: str = "127.0.0.1", port: int = 8770) -> int:
    import uvicorn

    app = create_app()
    print(f"KBondWatcher Admin http://{host}:{port}/")
    uvicorn.run(app, host=host, port=port, log_level="info")
    return 0
