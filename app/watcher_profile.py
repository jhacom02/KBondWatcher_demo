from __future__ import annotations

import os
import time
from typing import Optional

from app import ENGINE_VERSION
from app.adapter import config_from_profile, slot_from_profile
from app.audit import append_audit
from app.controller import resolve_yield_prefix
from app.defaults import DEFAULTS
from app.machine import load_or_create_machine
from app.profile import load_profile, prefs_snapshot
from app.runtime_status import RuntimeStatus, write_runtime_status
from core import (
    AppStatus,
    WatcherSession,
    evaluate,
    format_message,
    get_logger,
    pnl_outside_sanity_band,
    setup_logger,
)
from core.perf_log import append_sent, sent_perf_path
from core.trigger import format_looking_for_label
from excel import ExcelBridge, ExcelBridgeError, ExcelDisconnected, StopRequested
import send
from send import SendError
from source import SourceReaderError, create_source_reader


def _status(
    *,
    state: str,
    profile_version: int,
    instrument: str,
    looking_for: str,
    threshold: float,
    pid: Optional[int] = None,
    last_quote: Optional[str] = None,
    last_pnl: Optional[float] = None,
    last_action: Optional[str] = None,
    last_error: Optional[str] = None,
) -> None:
    write_runtime_status(
        RuntimeStatus(
            state=state,
            watcher_pid=pid or os.getpid(),
            instrument=instrument,
            looking_for=looking_for,
            threshold=threshold,
            last_quote=last_quote,
            last_pnl=last_pnl,
            last_action=last_action,
            last_error=last_error,
            profile_version=profile_version,
            engine_version=ENGINE_VERSION,
        )
    )


def run_watcher_from_profile() -> int:
    from main import (
        clear_pid_file,
        clear_stop_flag,
        collect_batch_matches,
        log_new_source_lines,
        pid_file_path,
        stop_requested,
        write_pid_file,
    )

    profile = load_profile()
    machine = load_or_create_machine()
    cfg = config_from_profile(profile, machine)
    setup_logger(cfg.log_path, level=cfg.log_level)
    log = get_logger()
    session = WatcherSession(status=AppStatus.STARTING)
    pid_path = pid_file_path(cfg.stop_flag_path)
    looking_label = format_looking_for_label(profile.instrument, profile.looking_for)
    threshold = float(profile.threshold)
    last_pnl: Optional[float] = None
    excel: Optional[ExcelBridge] = None
    trader_id = profile.trader_id or profile.profile_name

    def publish(state: str, **kwargs) -> None:
        _status(
            state=state,
            profile_version=profile.profile_version,
            instrument=profile.instrument,
            looking_for=profile.looking_for,
            threshold=threshold,
            **kwargs,
        )

    excel = ExcelBridge(
        workbook_name=cfg.excel_workbook,
        sheet_name=cfg.excel_sheet,
    )
    excel.set_stop_check(lambda: stop_requested(cfg.stop_flag_path))

    try:
        write_pid_file(pid_path)
        clear_stop_flag(cfg.stop_flag_path, required=True)
        publish("STARTING")

        prefix = 0.0
        excel.connect()
        prefix = resolve_yield_prefix(profile, excel)
        session.status = AppStatus.WATCHING
        publish("WATCHING", last_action="Start Successful")
        log.info(
            "WATCHING | mode=%s looking_for=%s threshold=%s profile_v=%s",
            cfg.mode,
            looking_label,
            threshold,
            profile.profile_version,
        )

        reader = create_source_reader(cfg)
        reader.find_source_window()
        send.ensure_target_window(cfg)
        reader.initialize_watermark(cfg.process_existing_on_start)
        poll_sec = cfg.poll_interval_ms / 1000.0
        slot = slot_from_profile(profile, prefix if prefix else float(profile.yield_prefix))

        while True:
            if stop_requested(cfg.stop_flag_path):
                session.status = AppStatus.STOPPED
                publish("STOPPED", last_action="Stopped")
                log.info("STOPPED")
                clear_stop_flag(cfg.stop_flag_path)
                return 0

            publish(session.status.value if hasattr(session.status, "value") else str(session.status))

            lines = reader.get_new_message_lines(
                process_existing_on_start=cfg.process_existing_on_start
            )
            if lines:
                try:
                    if stop_requested(cfg.stop_flag_path):
                        publish("STOPPED", last_action="Stopped")
                        return 0
                    log_new_source_lines(
                        log,
                        mode=cfg.mode,
                        looking_for=looking_label,
                        threshold=threshold,
                        lines=lines,
                    )
                    matches = collect_batch_matches(
                        lines,
                        [slot],
                        session.processed_fingerprints,
                        skip_fingerprints=cfg.mode == 3,
                    )
                    if matches:
                        line, quote, matched_slot, fp = matches[0]
                        if cfg.mode != 3:
                            session.processed_fingerprints.add(fp)
                        t0 = time.perf_counter()
                        session.status = AppStatus.QUOTE_FOUND
                        log.info(
                            "QUOTE_FOUND | %s | raw_token=%s | %.3f | %s | raw_line=%s",
                            quote.instrument,
                            quote.raw_token,
                            quote.yield_value,
                            quote.side,
                            line.watermark_key,
                        )
                        session.status = AppStatus.CALCULATING
                        t_excel = time.perf_counter()
                        pnl = excel.write_yield_read_pnl(
                            matched_slot.input_cell,
                            matched_slot.pnl_cell,
                            quote.yield_value,
                        )
                        excel_ms = (time.perf_counter() - t_excel) * 1000.0
                        last_pnl = pnl
                        if pnl_outside_sanity_band(
                            pnl, threshold, cfg.excel_pnl_sanity_band
                        ):
                            raise ExcelBridgeError(
                                f"PnL {pnl} outside sanity band "
                                f"{cfg.excel_pnl_sanity_band} of threshold {threshold}"
                            )
                        result = evaluate(
                            quote,
                            pnl,
                            threshold,
                            threshold_op=(profile.threshold_op or "<="),
                        )
                        append_audit(
                            "TRIGGER_RESULT",
                            {
                                "trader_id": trader_id,
                                "machine_id": machine.machine_id,
                                "engine_version": ENGINE_VERSION,
                                "quantity": quote.quantity,
                                "raw_line": line.watermark_key,
                                "parsed_yield": quote.yield_value,
                                "pnl": pnl,
                                "result": "TRIGGERED" if result.triggered else "NO_TRIGGER",
                                "error_message": result.reason,
                                **prefs_snapshot(profile),
                                "instrument": quote.instrument,
                                "looking_for": matched_slot.looking_for,
                                "threshold": threshold,
                            },
                        )
                        if not result.triggered:
                            session.status = AppStatus.WATCHING
                            log.info("NO_TRIGGER | %s", result.reason)
                            publish(
                                "WATCHING",
                                last_quote=f"{quote.instrument} {quote.raw_token}",
                                last_pnl=pnl,
                                last_action="Quote Skipped",
                            )
                        else:
                            log.info("TRIGGERED")
                            text = format_message(cfg.message_template, quote, pnl)
                            append_audit(
                                "SEND_ATTEMPT",
                                {
                                    "trader_id": trader_id,
                                    "machine_id": machine.machine_id,
                                    "engine_version": ENGINE_VERSION,
                                    "quantity": quote.quantity,
                                    "raw_line": line.watermark_key,
                                    "parsed_yield": quote.yield_value,
                                    "pnl": pnl,
                                    "sent_message": text,
                                    **prefs_snapshot(profile),
                                    "instrument": quote.instrument,
                                    "looking_for": matched_slot.looking_for,
                                    "threshold": threshold,
                                },
                            )
                            t_send = time.perf_counter()
                            try:
                                time.sleep(DEFAULTS.perf_pad_seconds)
                                send.send_text(text, cfg)
                                send_ok = True
                                send_err = None
                            except SendError as exc:
                                send_ok = False
                                send_err = str(exc)
                                append_audit(
                                    "SEND_RESULT",
                                    {
                                        "trader_id": trader_id,
                                        "machine_id": machine.machine_id,
                                        "engine_version": ENGINE_VERSION,
                                        "sent_message": text,
                                        "result": "FAILED",
                                        "error_message": send_err,
                                        "pnl": pnl,
                                        "raw_line": line.watermark_key,
                                        **prefs_snapshot(profile),
                                        "threshold": threshold,
                                    },
                                )
                                raise
                            pad_ms = DEFAULTS.perf_pad_seconds * 1000.0
                            send_ms = (time.perf_counter() - t_send) * 1000.0
                            total_ms = (time.perf_counter() - t0) * 1000.0
                            raw_for_log = (
                                line.watermark_key if cfg.mode == 3 else line.text
                            )
                            append_sent(
                                sent_perf_path(cfg.log_path),
                                mode=cfg.mode,
                                looking_for=looking_label,
                                raw_line=raw_for_log,
                                sent_message=text,
                                total_ms=max(0.0, total_ms - pad_ms),
                                excel_ms=excel_ms,
                                send_ms=max(0.0, send_ms - pad_ms),
                            )
                            append_audit(
                                "SEND_RESULT",
                                {
                                    "trader_id": trader_id,
                                    "machine_id": machine.machine_id,
                                    "engine_version": ENGINE_VERSION,
                                    "quantity": quote.quantity,
                                    "raw_line": raw_for_log,
                                    "parsed_yield": quote.yield_value,
                                    "pnl": pnl,
                                    "sent_message": text,
                                    "result": "SUCCESS",
                                    **prefs_snapshot(profile),
                                    "instrument": quote.instrument,
                                    "looking_for": matched_slot.looking_for,
                                    "threshold": threshold,
                                },
                            )
                            session.status = AppStatus.SENT
                            log.info("SENT")
                            publish(
                                "SENT",
                                last_quote=f"{quote.instrument} {quote.raw_token}",
                                last_pnl=pnl,
                                last_action=f"Message Sent: {text}",
                            )
                            if cfg.sent_after == "exit":
                                log.info("EXIT")
                                return 0
                            if cfg.mode != 3:
                                reader.reseed_watermark_from_visible()
                            session.status = AppStatus.WATCHING
                            publish("WATCHING")
                            log.info("WATCHING")
                except ExcelDisconnected as exc:
                    raise ExcelBridgeError(str(exc)) from exc

            time.sleep(poll_sec)

    except StopRequested:
        publish("STOPPED", last_action="Stopped")
        return 0
    except (ExcelBridgeError, SourceReaderError, SendError, ExcelDisconnected) as exc:
        log.error("ERROR | %s", exc)
        append_audit(
            "ERROR",
            {
                "trader_id": trader_id,
                "machine_id": machine.machine_id,
                "engine_version": ENGINE_VERSION,
                "error_message": str(exc),
                **prefs_snapshot(profile),
            },
        )
        publish("ERROR", last_error=str(exc), last_pnl=last_pnl)
        return 1
    except Exception as exc:
        log.exception("ERROR | unexpected: %s", exc)
        append_audit(
            "ERROR",
            {
                "trader_id": trader_id,
                "machine_id": machine.machine_id,
                "error_message": str(exc),
                **prefs_snapshot(profile),
            },
        )
        publish("ERROR", last_error=str(exc), last_pnl=last_pnl)
        return 1
    finally:
        clear_pid_file(pid_path)
        if excel is not None:
            excel.close()
