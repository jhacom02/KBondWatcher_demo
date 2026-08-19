from __future__ import annotations

import json
import os
import secrets
import threading
import time
from pathlib import Path
from typing import Any, Optional

from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from starlette.middleware.sessions import SessionMiddleware

from app import ENGINE_VERSION
from app.adapter import config_from_profile
from app.admin_client import (
    AdminClientError,
    fetch_signed_profile,
    submit_profile_draft,
    upload_audit_batch,
)
from app.audit import append_audit, iter_audit, read_audit_upload_status, write_audit_upload_status
from app.calibration import CalibrationError, capture_click_ratio
from app.controller import (
    ensure_device_activated,
    list_open_workbooks,
    refresh_lease_if_possible,
    start_watcher_subprocess,
    stop_watcher,
)
from app.crypto_sign import admin_sign_payload
from app.deploy_mode import is_dev, is_pilot
from app.license import (
    load_or_create_device,
    save_profile_signature,
)
from app.machine import load_or_create_machine, save_machine
from app.paths import audit_cursor_path, local_token_path
from app.policy_poll import get_last_profile_sync, start_policy_poller
from app.profile import (
    ProfileError,
    TraderProfile,
    apply_signed_profile,
    load_profile,
    load_profile_draft,
    save_profile,
    save_profile_draft,
)
from app.runtime_status import read_runtime_status

TEMPLATES = Jinja2Templates(directory=str(Path(__file__).parent / "templates"))


def _load_or_create_token() -> str:
    path = local_token_path()
    if path.is_file():
        return path.read_text(encoding="utf-8").strip()
    token = secrets.token_urlsafe(24)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(token, encoding="utf-8")
    return token


LOCAL_TOKEN = _load_or_create_token()


def create_app() -> FastAPI:
    app = FastAPI(docs_url=None, redoc_url=None)
    app.add_middleware(
        SessionMiddleware,
        secret_key=LOCAL_TOKEN,
        same_site="strict",
        https_only=False,
    )

    @app.middleware("http")
    async def security_headers(request: Request, call_next):
        host = (request.headers.get("host") or "").split(":")[0]
        if host not in {"127.0.0.1", "localhost"}:
            return JSONResponse({"error": "host not allowed"}, status_code=403)
        origin = request.headers.get("origin")
        if origin and not (
            origin.startswith("http://127.0.0.1") or origin.startswith("http://localhost")
        ):
            return JSONResponse({"error": "origin not allowed"}, status_code=403)
        if request.method not in {"GET", "HEAD", "OPTIONS"} and request.url.path.startswith("/api/"):
            token = request.headers.get("x-kbond-token") or request.cookies.get("kbond_token")
            if token != LOCAL_TOKEN:
                if not request.session.get("ok"):
                    return JSONResponse({"error": "unauthorized"}, status_code=401)
        response = await call_next(request)
        response.headers["Cache-Control"] = "no-store"
        return response

    @app.get("/", response_class=HTMLResponse)
    async def index(request: Request):
        ensure_device_activated()
        machine = load_or_create_machine()
        try:
            profile = load_profile()
        except ProfileError:
            try:
                profile = load_profile_draft()
            except ProfileError:
                profile = TraderProfile(profile_name="default", trader_id="trader")
        status = read_runtime_status()
        return TEMPLATES.TemplateResponse(
            request,
            "index.html",
            {
                "profile": profile,
                "machine": machine,
                "status": status,
                "engine_version": ENGINE_VERSION,
                "token": LOCAL_TOKEN,
                "workbooks": [],
                "message": request.query_params.get("msg", ""),
                "deploy_mode": "pilot" if is_pilot() else "dev",
                "audit_upload": read_audit_upload_status(),
            },
        )

    @app.get("/api/status")
    async def api_status():
        status = read_runtime_status()
        return {
            **status.to_dict(),
            "deploy_mode": "pilot" if is_pilot() else "dev",
            "audit_upload": read_audit_upload_status(),
            "profile_sync": get_last_profile_sync(),
        }

    @app.get("/api/workbooks")
    async def api_workbooks():
        return {"workbooks": list_open_workbooks()}

    def _require_stopped() -> None:
        st = read_runtime_status().state
        if st not in {"STOPPED", "ERROR", ""}:
            raise HTTPException(400, detail="stop watcher before changing profile")

    @app.post("/api/profile")
    async def api_profile_save(request: Request):
        """Save local draft. Pilot does not auto-apply; use submit + apply."""
        _require_stopped()
        data = await request.json()
        try:
            existing = load_profile()
            version = existing.profile_version
        except ProfileError:
            version = 0
        profile = TraderProfile.from_dict({**data, "profile_version": version})
        try:
            draft = save_profile_draft(profile)
        except ProfileError as exc:
            raise HTTPException(400, str(exc)) from exc

        machine = load_or_create_machine()
        append_audit(
            "PROFILE_SAVED",
            {
                "trader_id": draft.trader_id or draft.profile_name,
                "machine_id": machine.machine_id,
                "profile_version": draft.profile_version,
                "engine_version": ENGINE_VERSION,
                "draft": True,
            },
        )

        # Dev without Admin: allow local activate + sign for smoke.
        admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
        if is_dev() and not admin_url:
            saved = save_profile(draft)
            sig = admin_sign_payload(saved.to_dict())
            save_profile_signature(sig)
            try:
                refresh_lease_if_possible(saved)
            except Exception:
                pass
            return {"ok": True, "profile": saved.to_dict(), "mode": "dev_local_apply"}

        return {"ok": True, "draft": draft.to_dict(), "mode": "draft"}

    @app.post("/api/profile/submit")
    async def api_profile_submit():
        _require_stopped()
        admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
        if not admin_url:
            raise HTTPException(400, "KBOND_ADMIN_URL required")
        try:
            draft = load_profile_draft()
        except ProfileError as exc:
            raise HTTPException(400, f"no draft: {exc}") from exc
        machine = load_or_create_machine()
        device = load_or_create_device(machine.machine_id)
        try:
            resp = submit_profile_draft(admin_url, device, draft)
        except AdminClientError as exc:
            append_audit(
                "PROFILE_REJECTED",
                {"error_message": str(exc), "trader_id": draft.trader_id},
            )
            raise HTTPException(400, str(exc)) from exc
        return {"ok": True, **resp}

    @app.post("/api/profile/apply")
    async def api_profile_apply():
        _require_stopped()
        admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
        if not admin_url:
            raise HTTPException(400, "KBOND_ADMIN_URL required")
        try:
            draft = load_profile_draft()
            trader_id = draft.trader_id or draft.profile_name
        except ProfileError:
            try:
                active = load_profile()
                trader_id = active.trader_id or active.profile_name
            except ProfileError as exc:
                raise HTTPException(400, str(exc)) from exc
        machine = load_or_create_machine()
        device = load_or_create_device(machine.machine_id)
        try:
            data = fetch_signed_profile(admin_url, device, trader_id)
            profile = TraderProfile.from_dict(data["profile"])
            apply_signed_profile(profile, str(data["signature"]))
            refresh_lease_if_possible(profile)
        except (AdminClientError, ProfileError, Exception) as exc:
            append_audit(
                "PROFILE_REJECTED",
                {"error_message": str(exc), "trader_id": trader_id},
            )
            raise HTTPException(400, str(exc)) from exc
        return {"ok": True, "profile": profile.to_dict()}

    @app.post("/api/start")
    async def api_start():
        result = start_watcher_subprocess()
        if not result.ok:
            raise HTTPException(400, result.message)
        return {"ok": True, "message": result.message}

    @app.post("/api/stop")
    async def api_stop():
        result = stop_watcher()
        if not result.ok:
            raise HTTPException(400, result.message)
        return {"ok": True, "message": result.message}

    @app.post("/api/calibrate")
    async def api_calibrate():
        _require_stopped()
        try:
            profile = load_profile()
        except ProfileError as exc:
            raise HTTPException(400, str(exc)) from exc
        machine = load_or_create_machine()
        cfg = config_from_profile(profile, machine)
        try:
            x, y = capture_click_ratio(cfg)
        except CalibrationError as exc:
            raise HTTPException(400, str(exc)) from exc
        machine.send_input_x = x
        machine.send_input_y = y
        save_machine(machine)
        return {"ok": True, "send_input_x": x, "send_input_y": y}

    @app.post("/api/test-send-target")
    async def api_test_send_target():
        from send import diagnose

        profile = load_profile()
        machine = load_or_create_machine()
        cfg = config_from_profile(profile, machine)
        return {"ok": True, "diagnose": diagnose(cfg)}

    @app.post("/api/test-send-message")
    async def api_test_send_message(request: Request):
        body = await request.json()
        if not body.get("confirm"):
            raise HTTPException(400, "confirmation required")
        from core.models import Quote
        from core.trigger import format_message
        import send as send_mod

        profile = load_profile()
        machine = load_or_create_machine()
        cfg = config_from_profile(profile, machine)
        sample = Quote(
            instrument=profile.instrument or "25-10",
            raw_line="test",
            raw_token="00+",
            yield_value=0.0,
            side="BUY",
        )
        text = format_message(cfg.message_template, sample, 0.0)
        send_mod.send_text(text, cfg)
        return {"ok": True, "sent": text}

    def _uploader_loop() -> None:
        cursor: Optional[str] = None
        marker = audit_cursor_path()
        if marker.is_file():
            cursor = marker.read_text(encoding="utf-8").strip() or None
        backoff = 10.0
        while True:
            admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
            if admin_url:
                rows = list(iter_audit(after_id=cursor))
                if rows:
                    try:
                        machine = load_or_create_machine()
                        device = load_or_create_device(machine.machine_id)
                        result = upload_audit_batch(admin_url, device, rows)
                        cursor = rows[-1].get("event_id") or cursor
                        marker.parent.mkdir(parents=True, exist_ok=True)
                        marker.write_text(cursor or "", encoding="utf-8")
                        write_audit_upload_status(
                            {
                                "ok": True,
                                "last_upload_at": time.time(),
                                "last_event_id": cursor,
                                "accepted": result.get("accepted"),
                                "duplicates": result.get("duplicates"),
                                "pending": 0,
                                "error": "",
                            }
                        )
                        backoff = 10.0
                    except Exception as exc:
                        write_audit_upload_status(
                            {
                                "ok": False,
                                "last_attempt_at": time.time(),
                                "last_event_id": cursor,
                                "pending": len(rows),
                                "error": str(exc),
                                "stale": True,
                            }
                        )
                        backoff = min(backoff * 1.5, 120.0)
                else:
                    write_audit_upload_status(
                        {
                            "ok": True,
                            "last_event_id": cursor,
                            "pending": 0,
                            "stale": False,
                            "error": "",
                        }
                    )
            time.sleep(backoff)

    @app.on_event("startup")
    def _startup() -> None:
        ensure_device_activated()
        start_policy_poller()
        thread = threading.Thread(target=_uploader_loop, name="audit-uploader", daemon=True)
        thread.start()

    return app


def run_local_web(host: str = "127.0.0.1", port: int = 8765) -> int:
    import uvicorn

    ensure_device_activated()
    app = create_app()
    print(f"KBondWatcher local UI http://{host}:{port}/")
    print(f"local token (header X-KBond-Token): {LOCAL_TOKEN}")
    uvicorn.run(app, host=host, port=port, log_level="info")
    return 0
