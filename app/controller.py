from __future__ import annotations

import os
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Optional

import psutil

from app import ENGINE_VERSION
from app.adapter import config_from_profile
from app.audit import append_audit
from app.defaults import DEFAULTS
from app.demo_expiry import check_demo_expiry
from app.deploy_mode import is_dev
from app.license import (
    LicenseError,
    activate_device,
    issue_lease,
    load_lease,
    load_or_create_device,
    load_profile_signature,
    verify_lease_for_start,
    verify_signed_profile,
)
from app.machine import load_or_create_machine
from app.paths import stop_flag_path
from app.profile import ProfileError, TraderProfile, load_profile
from app.runtime_status import RuntimeStatus, read_runtime_status, write_runtime_status
from excel import ExcelBridge, ExcelBridgeError, ExcelDisconnected, bind_open_workbook
from excel.bridge import workbook_identity
from send import ensure_target_window
from source import create_source_reader

_watcher_job = None


def _ensure_watcher_job() -> Any:
    """Windows job so watcher dies when the serve process exits (console close)."""
    global _watcher_job
    if sys.platform != "win32":
        return None
    if _watcher_job is not None:
        return _watcher_job
    import win32job

    job = win32job.CreateJobObject(None, "")
    info = win32job.QueryInformationJobObject(job, win32job.JobObjectExtendedLimitInformation)
    info["BasicLimitInformation"]["LimitFlags"] |= win32job.JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE
    win32job.SetInformationJobObject(job, win32job.JobObjectExtendedLimitInformation, info)
    _watcher_job = job
    return job


def _assign_watcher_to_job(pid: int) -> bool:
    job = _ensure_watcher_job()
    if job is None:
        return False
    import win32api
    import win32con
    import win32job

    access = win32con.PROCESS_SET_QUOTA | win32con.PROCESS_TERMINATE
    handle = win32api.OpenProcess(access, False, int(pid))
    try:
        win32job.AssignProcessToJobObject(job, handle)
    finally:
        win32api.CloseHandle(handle)
    return True


@dataclass
class PreflightResult:
    ok: bool
    message: str = ""


def _pid_alive(pid: Optional[int]) -> bool:
    if not pid or pid <= 0:
        return False
    try:
        return psutil.pid_exists(int(pid))
    except Exception:
        return False


def resolve_yield_prefix(profile: TraderProfile, excel: ExcelBridge) -> float:
    return float(profile.yield_prefix)


def preflight(profile: TraderProfile) -> PreflightResult:
    try:
        profile.validate()
    except ProfileError as exc:
        return PreflightResult(False, str(exc))

    machine = load_or_create_machine()
    device = load_or_create_device(machine.machine_id)
    try:
        check_demo_expiry()
        verify_signed_profile(profile, load_profile_signature())
        lease = load_lease()
        admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
        if admin_url:
            try:
                from app.admin_client import refresh_lease_from_admin

                lease = refresh_lease_from_admin(admin_url, device, profile)
            except Exception:
                # Use cached lease when Admin is unreachable
                pass
        if lease is None and is_dev():
            lease = issue_lease(device=device, profile=profile, ttl_seconds=3600)
        verify_lease_for_start(device=device, profile=profile, lease=lease)
    except LicenseError as exc:
        append_audit(
            "LICENSE_REJECTED",
            {
                "trader_id": profile.trader_id or profile.profile_name,
                "machine_id": machine.machine_id,
                "error_message": str(exc),
                "engine_version": ENGINE_VERSION,
                "profile_version": profile.profile_version,
            },
        )
        return PreflightResult(False, str(exc))

    cfg = config_from_profile(profile, machine)
    excel = ExcelBridge(
        workbook_name=cfg.excel_workbook,
        sheet_name=cfg.excel_sheet,
    )
    try:
        wb = bind_open_workbook(profile.excel_workbook)
        name, full = workbook_identity(wb)
        cfg_norm = profile.excel_workbook.replace("/", "\\").lower()
        full_norm = full.replace("/", "\\").lower()
        if Path(full_norm).name != Path(cfg_norm).name and full_norm != cfg_norm:
            if full_norm != cfg_norm:
                try:
                    if str(Path(profile.excel_workbook).expanduser().resolve()).replace("/", "\\").lower() != full_norm:
                        return PreflightResult(
                            False,
                            f"workbook FullName mismatch profile={profile.excel_workbook!r} open={full!r}",
                        )
                except OSError:
                    return PreflightResult(
                        False,
                        f"workbook FullName mismatch profile={profile.excel_workbook!r} open={full!r}",
                    )
        excel.connect()
        prefix = resolve_yield_prefix(profile, excel)
        excel.read_cell_text(profile.yield_input_cell)
        excel.read_cell_text(profile.pnl_cell)
        _ = prefix
    except (ExcelBridgeError, ExcelDisconnected, OSError, ValueError) as exc:
        return PreflightResult(False, f"excel preflight failed: {exc}")
    finally:
        try:
            excel.close()
        except Exception:
            pass

    try:
        reader = create_source_reader(cfg)
        reader.find_source_window()
    except Exception as exc:
        return PreflightResult(False, f"source preflight failed: {exc}")

    try:
        ensure_target_window(cfg)
    except Exception as exc:
        return PreflightResult(False, f"send target preflight failed: {exc}")

    return PreflightResult(True, "ok")


def list_open_workbooks() -> list[dict[str, Any]]:
    """UI convenience only — does not replace runtime GetObject binding."""
    out: list[dict[str, Any]] = []
    try:
        import pythoncom
        import win32com.client
    except ImportError:
        return out
    pythoncom.CoInitialize()
    try:
        try:
            app = win32com.client.GetObject(Class="Excel.Application")
        except Exception:
            return out
        for i in range(1, int(app.Workbooks.Count) + 1):
            wb = app.Workbooks(i)
            try:
                full = str(wb.FullName)
                name = str(wb.Name)
            except Exception:
                continue
            sheets = []
            try:
                for s in range(1, int(wb.Worksheets.Count) + 1):
                    sheets.append(str(wb.Worksheets(s).Name))
            except Exception:
                sheets = []
            out.append({"name": name, "full_name": full, "sheets": sheets})
    finally:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return out


def start_watcher_subprocess(python_exe: Optional[str] = None) -> PreflightResult:
    ensure_device_activated()
    status = read_runtime_status()
    if status.state not in {"STOPPED", "ERROR", "SENT", "DONE"} and _pid_alive(status.watcher_pid):
        return PreflightResult(False, f"watcher already running pid={status.watcher_pid}")
    if status.watcher_pid and not _pid_alive(status.watcher_pid):
        write_runtime_status(
            RuntimeStatus(state="STOPPED", engine_version=ENGINE_VERSION, last_action="stale pid cleared")
        )

    try:
        profile = load_profile()
    except ProfileError as exc:
        return PreflightResult(False, str(exc))

    check = preflight(profile)
    if not check.ok:
        write_runtime_status(
            RuntimeStatus(
                state="ERROR",
                last_error=check.message,
                engine_version=ENGINE_VERSION,
                profile_version=profile.profile_version,
                instrument=profile.instrument,
                looking_for=profile.looking_for,
                threshold=float(profile.threshold),
            )
        )
        return check

    machine = load_or_create_machine()
    flag = stop_flag_path()
    if flag.is_file():
        try:
            flag.unlink()
        except OSError:
            pass

    write_runtime_status(
        RuntimeStatus(
            state="STARTING",
            engine_version=ENGINE_VERSION,
            profile_version=profile.profile_version,
            instrument=profile.instrument,
            looking_for=profile.looking_for,
            threshold=float(profile.threshold),
        )
    )

    exe = python_exe or sys.executable
    if getattr(sys, "frozen", False):
        cmd = [exe, "--run-profile"]
    else:
        cmd = [exe, str(Path(__file__).resolve().parents[1] / "main.py"), "--run-profile"]
    creationflags = 0
    if sys.platform == "win32":
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
    proc = subprocess.Popen(
        cmd,
        cwd=str(Path(__file__).resolve().parents[1]),
        creationflags=creationflags,
    )
    job_ok = False
    try:
        job_ok = _assign_watcher_to_job(proc.pid)
    except Exception:
        job_ok = False
    append_audit(
        "WATCHER_STARTED",
        {
            "trader_id": profile.trader_id or profile.profile_name,
            "machine_id": machine.machine_id,
            "engine_version": ENGINE_VERSION,
            "profile_version": profile.profile_version,
            "watcher_pid": proc.pid,
            "instrument": profile.instrument,
            "looking_for": profile.looking_for,
            "quantity": profile.required_qty,
            "threshold": profile.threshold,
        },
    )
    msg = f"started pid={proc.pid}"
    if not job_ok:
        msg += " · console-close kill unavailable"
    return PreflightResult(True, msg)


def stop_watcher(soft_wait_seconds: Optional[float] = None) -> PreflightResult:
    wait_s = float(soft_wait_seconds if soft_wait_seconds is not None else DEFAULTS.stop_soft_wait_seconds)
    status = read_runtime_status()
    pid = status.watcher_pid
    write_runtime_status(
        RuntimeStatus(
            state="STOPPING",
            watcher_pid=pid,
            engine_version=ENGINE_VERSION,
            profile_version=status.profile_version,
            instrument=status.instrument,
            looking_for=status.looking_for,
            threshold=status.threshold,
            last_action="stop requested",
        )
    )
    flag = stop_flag_path()
    flag.parent.mkdir(parents=True, exist_ok=True)
    flag.write_text("stop\n", encoding="utf-8")

    deadline = time.monotonic() + wait_s
    while time.monotonic() < deadline:
        if not _pid_alive(pid):
            break
        time.sleep(0.2)

    if _pid_alive(pid):
        try:
            proc = psutil.Process(int(pid))
            proc.terminate()
            try:
                proc.wait(timeout=3)
            except psutil.TimeoutExpired:
                proc.kill()
        except (psutil.Error, ValueError) as exc:
            return PreflightResult(False, f"failed to terminate pid={pid}: {exc}")

    write_runtime_status(
        RuntimeStatus(state="STOPPED", engine_version=ENGINE_VERSION, last_action="stopped")
    )
    append_audit(
        "WATCHER_STOPPED",
        {
            "watcher_pid": pid,
            "engine_version": ENGINE_VERSION,
            "profile_version": status.profile_version,
        },
    )
    return PreflightResult(True, "stopped")


def ensure_device_activated() -> dict[str, Any]:
    machine = load_or_create_machine()
    device = load_or_create_device(machine.machine_id)
    if device.activated:
        return device.to_dict()

    admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
    if admin_url:
        try:
            from app.admin_client import register_device

            try:
                profile = load_profile()
                trader_id = profile.trader_id or profile.profile_name
            except ProfileError:
                trader_id = "pending"
            resp = register_device(admin_url, device, trader_id)
            if resp.get("activated"):
                activate_device(device)
                device = load_or_create_device(machine.machine_id)
        except Exception:
            pass

    if not device.activated and is_dev():
        activate_device(device)
        device = load_or_create_device(machine.machine_id)
    return device.to_dict()


def refresh_lease_if_possible(profile: TraderProfile) -> None:
    """Best-effort lease refresh — never called from send path."""
    machine = load_or_create_machine()
    device = load_or_create_device(machine.machine_id)
    if not device.activated or device.disabled:
        return
    admin_url = (os.environ.get("KBOND_ADMIN_URL") or "").strip()
    if not admin_url:
        if is_dev():
            issue_lease(device=device, profile=profile, ttl_seconds=3600)
        return
    try:
        from app.admin_client import refresh_lease_from_admin

        refresh_lease_from_admin(admin_url, device, profile)
    except Exception:
        return
