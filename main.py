from __future__ import annotations

import argparse
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

from config import Config, ConfigError
from excel import (
    ExcelBridge,
    ExcelBridgeError,
    ExcelDisconnected,
    InstrumentSlot,
    StopRequested,
    bind_open_workbook,
    is_excel_gone,
)
from source import SourceLine, SourceReaderError, create_source_reader, format_parser_result, message_fingerprint, parse_quote_line
from core import (
    AppStatus,
    Quote,
    WatcherSession,
    evaluate,
    format_message,
    get_logger,
    pnl_outside_sanity_band,
    setup_logger,
)
from core.perf_log import append_sent, sent_perf_path, summarize as summarize_sent_perf
import send
from send import SendError


def pid_file_path(stop_flag_path: Path) -> Path:
    return stop_flag_path.with_suffix(".pid")


def write_pid_file(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(str(os.getpid()), encoding="ascii")


def clear_pid_file(path: Path) -> None:
    try:
        if not path.is_file():
            return
        stored = path.read_text(encoding="ascii").strip()
        if stored == str(os.getpid()):
            path.unlink()
    except OSError as exc:
        get_logger().error("Failed to clear pid file %s: %s", path, exc)


def clear_stop_flag(path: Path, *, required: bool = False) -> None:
    try:
        if path.is_file():
            path.unlink()
    except OSError as exc:
        if required:
            raise
        get_logger().error("Failed to clear stop flag %s: %s", path, exc)


def stop_requested(path: Path) -> bool:
    return path.is_file()


def _env_file_values(path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    for line in path.read_text(encoding="utf-8").splitlines():
        stripped = line.strip()
        if not stripped or stripped.startswith("#") or "=" not in stripped:
            continue
        key, _, val = stripped.partition("=")
        values[key.strip()] = val.strip().strip('"').strip("'")
    return values


def report_startup_error(config_path: str | Path, message: str) -> None:
    path = Path(config_path).expanduser()
    if not path.is_file():
        get_logger().error("Failed to write startup Excel error: config not found")
        return
    try:
        values = _env_file_values(path)
    except OSError as exc:
        get_logger().error("Failed to read config for Excel error: %s", exc)
        return
    workbook = (values.get("EXCEL_WORKBOOK") or "").strip()
    sheet_name = (values.get("EXCEL_SHEET") or "").strip()
    status_cell = (values.get("EXCEL_STATUS_CELL") or "").strip()
    action_cell = (values.get("EXCEL_LAST_ACTION_CELL") or "").strip()
    if not workbook or not status_cell or not action_cell:
        get_logger().error(
            "Failed to write startup Excel error: missing workbook/status/J2 in .env"
        )
        return
    pythoncom = None
    try:
        import pythoncom as _pythoncom

        pythoncom = _pythoncom
        pythoncom.CoInitialize()
        wb = bind_open_workbook(workbook)
        ws = wb.Worksheets(sheet_name) if sheet_name else wb.ActiveSheet
        ws.Range(status_cell).Value = AppStatus.ERROR.value
        ws.Range(action_cell).Value = format_last_action(f"Error: {message[:200]}")
    except Exception as exc:
        get_logger().error("Failed to write startup Excel error: %s", exc)
    finally:
        if pythoncom is not None:
            try:
                pythoncom.CoUninitialize()
            except Exception as exc:
                get_logger().error("Excel CoUninitialize failed: %s", exc)


def format_last_action(action: str) -> str:
    stamp = datetime.now().strftime("%H:%M:%S")
    return f"({stamp}) {action}"


def _write_error_cells(
    excel: Optional[ExcelBridge],
    cfg: Config,
    message: str,
    last_pnl: Optional[float] = None,
) -> None:
    action = format_last_action(f"Error: {message[:200]}")
    target = excel
    created = False
    try:
        if target is None:
            target = _build_excel(cfg)
            target.connect()
            created = True
        target.update_status(
            AppStatus.ERROR,
            last_pnl=last_pnl,
            last_action=action,
            ignore_error=True,
        )
    except Exception as exc:
        get_logger().error("Failed to write Excel error: %s", exc)
    finally:
        if created and target is not None:
            target.close()


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=".env")
    parser.add_argument("--diagnose-source", action="store_true")
    parser.add_argument("--diagnose-send", action="store_true")
    parser.add_argument("--test-send", action="store_true")
    parser.add_argument("--test-parser", metavar="LINE")
    parser.add_argument("--perf-summary", action="store_true")
    parser.add_argument("--run-profile", action="store_true")
    parser.add_argument("--serve", action="store_true")
    parser.add_argument("--serve-admin", action="store_true")
    return parser


def _build_excel(cfg: Config) -> ExcelBridge:
    return ExcelBridge(
        workbook_name=cfg.excel_workbook,
        sheet_name=cfg.excel_sheet,
        status_cell=cfg.excel_status_cell,
        looking_for_cell=cfg.excel_looking_for_cell,
        last_quote_cell=cfg.excel_last_quote_cell,
        last_pnl_cell=cfg.excel_last_pnl_cell,
        last_action_cell=cfg.excel_last_action_cell,
        watch_cell=cfg.excel_watch_cell,
        pnl_threshold_cell=cfg.excel_pnl_threshold_cell,
        slot_rows=cfg.excel_slot_rows,
        rows_10y=cfg.excel_rows_10y,
        rows_3y=cfg.excel_rows_3y,
        prefix_3y_cell=cfg.excel_prefix_3y_cell,
        prefix_10y_cell=cfg.excel_prefix_10y_cell,
        instrument_col=cfg.excel_instrument_col,
        qty_col=cfg.excel_qty_col,
        input_col=cfg.excel_input_col,
        pnl_col=cfg.excel_pnl_col,
        pnl_row_offset=cfg.excel_pnl_row_offset,
    )


def run_diagnose_source(cfg: Config) -> int:
    reader = create_source_reader(cfg)
    print(reader.diagnose(max_messages=200))
    return 0


def run_diagnose_send(cfg: Config) -> int:
    print(send.diagnose(cfg))
    return 0


def run_perf_summary(cfg: Config) -> int:
    text = summarize_sent_perf(sent_perf_path(cfg.log_path))
    print(text)
    if text.startswith("no sent perf"):
        return 1
    return 0


def run_test_send(cfg: Config) -> int:
    sample = Quote(
        instrument="25-10",
        raw_line="25-10 test",
        raw_token="00+",
        yield_value=3.0,
        side="BUY",
    )
    text = format_message(cfg.message_template, sample, pnl=0.0)
    send.send_text(text, cfg)
    print(f"sent to {cfg.send_window_title!r}: {text}")
    return 0


def run_test_parser(cfg: Config, line: str) -> int:
    excel = _build_excel(cfg)
    try:
        excel.connect()
        slots, looking_for, _threshold = excel.load_slots()
    finally:
        excel.close()
    for slot in slots:
        quote = parse_quote_line(
            line=line,
            target=slot.instrument,
            yield_prefix=slot.yield_prefix,
            required_side=slot.required_side,
            required_qty=slot.qty_abs,
        )
        if quote is None:
            continue
        print(f"looking_for={looking_for} slot_row={slot.row}")
        print(format_parser_result(quote))
        return 0
    instruments = ", ".join(s.instrument for s in slots)
    print(f"No match for active slots [{instruments}] in line: {line!r}")
    return 1


def _match_quote(
    line: str,
    slots: list[InstrumentSlot],
) -> tuple[Optional[Quote], Optional[InstrumentSlot]]:
    for slot in slots:
        quote = parse_quote_line(
            line=line,
            target=slot.instrument,
            yield_prefix=slot.yield_prefix,
            required_side=slot.required_side,
            required_qty=slot.qty_abs,
        )
        if quote is not None:
            return quote, slot
    return None, None


def watch_identity(
    slot: InstrumentSlot, threshold: float
) -> tuple[str, str, int, float]:
    return (slot.instrument, slot.looking_for, slot.qty_abs, threshold)


def collect_batch_matches(
    lines: list[SourceLine],
    slots: list[InstrumentSlot],
    processed_fps: set[str],
    *,
    skip_fingerprints: bool = False,
) -> list[tuple[SourceLine, Quote, InstrumentSlot, str]]:
    matches: list[tuple[SourceLine, Quote, InstrumentSlot, str]] = []
    for line in lines:
        quote, slot = _match_quote(line.text, slots)
        if quote is None or slot is None:
            continue
        fp = message_fingerprint(line.watermark_key)
        if not skip_fingerprints and fp in processed_fps:
            continue
        matches.append((line, quote, slot, fp))
    if len(matches) >= 2:
        raise SourceReaderError(f"ambiguous quotes in one poll: {len(matches)}")
    return matches


LINE_LOG_MAX_CHARS = 160
LINE_LOG_MAX_PER_POLL = 20


def _truncate_log_text(text: str, max_chars: int = LINE_LOG_MAX_CHARS) -> str:
    value = text or ""
    if len(value) <= max_chars:
        return value
    if max_chars <= 1:
        return "…"
    return value[: max_chars - 1] + "…"


def log_new_source_lines(
    log,
    *,
    mode: int,
    looking_for: str,
    threshold: float,
    lines: list[SourceLine],
) -> None:
    limit = LINE_LOG_MAX_PER_POLL
    for item in lines[:limit]:
        raw = item.watermark_key if mode == 3 else item.text
        log.info(
            "LINE | mode=%s | looking_for=%s | threshold=%s | raw_line=%s",
            mode,
            looking_for,
            int(threshold),
            _truncate_log_text(raw),
        )
    omitted = len(lines) - limit
    if omitted > 0:
        log.info("LINE_OMITTED | +%s", omitted)


def excel_failure_action(exc: BaseException, *, calculating: bool) -> str:
    if isinstance(exc, StopRequested):
        return "stop"
    gone = isinstance(exc, ExcelDisconnected) or is_excel_gone(exc)
    if gone and calculating:
        return "error"
    if gone:
        return "wait"
    return "error"


def run_watcher(cfg: Config) -> int:
    log = get_logger()
    session = WatcherSession(status=AppStatus.STARTING)
    excel: Optional[ExcelBridge] = None
    looking_for: Optional[str] = None
    threshold: Optional[float] = None
    last_pnl: Optional[float] = None
    pid_path = pid_file_path(cfg.stop_flag_path)
    workbook_file = Path(cfg.excel_workbook).name

    def _apply_stopped(looking_for: Optional[str] = None) -> None:
        session.status = AppStatus.STOPPED
        log.info("STOPPED")
        if excel is not None:
            try:
                excel.update_status(
                    AppStatus.STOPPED,
                    looking_for=looking_for,
                    last_action=format_last_action("Stopped"),
                )
            except ExcelDisconnected as exc:
                log.info("Excel status update failed: %s", exc)
        clear_stop_flag(cfg.stop_flag_path)

    def _enter_excel_wait(reason: BaseException) -> None:
        session.status = AppStatus.EXCEL_WAIT
        log.info("EXCEL_WAIT | %s", reason)
        if excel is None:
            return
        excel.release_workbook()
        try:
            excel.update_status(
                AppStatus.EXCEL_WAIT,
                looking_for=looking_for,
                last_action=format_last_action(
                    f"Excel closed; waiting to reopen {workbook_file}"
                ),
            )
        except Exception as exc:
            log.info("Excel status update failed: %s", exc)

    def _resume_excel() -> None:
        nonlocal looking_for, threshold
        assert excel is not None
        excel.connect()
        slots_now, looking_for, threshold = excel.load_slots()
        excel.update_status(
            AppStatus.WATCHING,
            looking_for=looking_for,
            last_action=format_last_action("Excel reconnected"),
        )
        session.status = AppStatus.WATCHING
        log.info(
            "WATCHING | mode=%s source=%s/%s send=%s/%s looking_for=%s "
            "threshold=%s slots=%s",
            cfg.mode,
            cfg.source_process_name or "(uia)",
            cfg.source_window_title,
            cfg.send_process_name,
            cfg.send_window_title,
            looking_for,
            threshold,
            [(s.instrument, s.row, s.required_side) for s in slots_now],
        )

    try:
        write_pid_file(pid_path)
        clear_stop_flag(cfg.stop_flag_path, required=True)
        excel = _build_excel(cfg)
        excel.set_stop_check(lambda: stop_requested(cfg.stop_flag_path))

        reader = create_source_reader(cfg)
        reader.find_source_window()
        send.ensure_target_window(cfg)
        reader.initialize_watermark(cfg.process_existing_on_start)
        poll_sec = cfg.poll_interval_ms / 1000.0

        try:
            excel.connect()
            slots, looking_for, threshold = excel.load_slots()
            excel.update_status(
                AppStatus.WATCHING,
                looking_for=looking_for,
                last_action=format_last_action("Start Successful"),
            )
            session.status = AppStatus.WATCHING
            log.info(
                "WATCHING | mode=%s source=%s/%s send=%s/%s looking_for=%s "
                "threshold=%s slots=%s",
                cfg.mode,
                cfg.source_process_name or "(uia)",
                cfg.source_window_title,
                cfg.send_process_name,
                cfg.send_window_title,
                looking_for,
                threshold,
                [(s.instrument, s.row, s.required_side) for s in slots],
            )
        except ExcelDisconnected as exc:
            _enter_excel_wait(exc)

        while True:
            if stop_requested(cfg.stop_flag_path):
                _apply_stopped(looking_for)
                return 0

            if session.status == AppStatus.EXCEL_WAIT:
                try:
                    _resume_excel()
                except ExcelDisconnected as exc:
                    log.info("EXCEL_WAIT | %s", exc)
                    time.sleep(poll_sec)
                continue

            lines = reader.get_new_message_lines(
                process_existing_on_start=cfg.process_existing_on_start
            )
            if lines:
                try:
                    slots, looking_for, threshold = excel.load_slots()
                    if stop_requested(cfg.stop_flag_path):
                        _apply_stopped(looking_for)
                        return 0
                    log_new_source_lines(
                        log,
                        mode=cfg.mode,
                        looking_for=looking_for,
                        threshold=threshold,
                        lines=lines,
                    )
                    matches = collect_batch_matches(
                        lines,
                        slots,
                        session.processed_fingerprints,
                        skip_fingerprints=cfg.mode == 3,
                    )
                    if matches:
                        line, quote, slot, fp = matches[0]
                        if cfg.mode != 3:
                            session.processed_fingerprints.add(fp)
                        t0 = time.perf_counter()

                        session.status = AppStatus.QUOTE_FOUND
                        log.info(
                            "QUOTE_FOUND | %s | raw_token=%s | %.3f | %s | row=%s | raw_line=%s",
                            quote.instrument,
                            quote.raw_token,
                            quote.yield_value,
                            quote.side,
                            slot.row,
                            line.watermark_key,
                        )

                        session.status = AppStatus.CALCULATING
                        t_excel = time.perf_counter()
                        pnl = excel.write_yield_read_pnl(
                            slot.input_cell,
                            slot.pnl_cell,
                            quote.yield_value,
                        )
                        excel_ms = (time.perf_counter() - t_excel) * 1000.0
                        last_pnl = pnl
                        if threshold is None:
                            raise ExcelBridgeError("PnL threshold is not loaded")
                        slots_now, looking_for, threshold_now = excel.load_slots()
                        before = watch_identity(slot, threshold)
                        after = watch_identity(slots_now[0], threshold_now)
                        if before != after:
                            raise ExcelBridgeError(
                                f"watch slot changed during PnL ({before} -> {after})"
                            )
                        slot = slots_now[0]
                        threshold = threshold_now
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
                            looking_for=slot.looking_for,
                        )

                        if not result.triggered:
                            session.status = AppStatus.NO_TRIGGER
                            log.info("NO_TRIGGER | %s", result.reason)
                            session.status = AppStatus.WATCHING
                            excel.update_status(
                                AppStatus.WATCHING,
                                looking_for=looking_for,
                                last_quote=f"{quote.instrument} {quote.raw_token}",
                                last_pnl=pnl,
                                last_action=format_last_action("Quote Skipped"),
                            )
                            log.info("WATCHING")
                        else:
                            session.status = AppStatus.TRIGGERED
                            log.info("TRIGGERED")
                            text = format_message(cfg.message_template, quote, pnl)
                            session.status = AppStatus.SENDING
                            t_send = time.perf_counter()
                            send.send_text(text, cfg)
                            send_ms = (time.perf_counter() - t_send) * 1000.0
                            total_ms = (time.perf_counter() - t0) * 1000.0
                            raw_for_log = (
                                line.watermark_key if cfg.mode == 3 else line.text
                            )
                            append_sent(
                                sent_perf_path(cfg.log_path),
                                mode=cfg.mode,
                                looking_for=looking_for or "",
                                raw_line=raw_for_log,
                                sent_message=text,
                                total_ms=total_ms,
                                excel_ms=excel_ms,
                                send_ms=send_ms,
                            )

                            session.status = AppStatus.SENT
                            log.info("SENT")
                            excel.update_status(
                                AppStatus.SENT,
                                looking_for=looking_for,
                                last_quote=f"{quote.instrument} {quote.raw_token}",
                                last_pnl=pnl,
                                last_action=format_last_action(f"Message Sent: {text}"),
                            )
                            if cfg.sent_after == "exit":
                                log.info("EXIT")
                                return 0
                            if cfg.mode != 3:
                                reader.reseed_watermark_from_visible()
                            session.status = AppStatus.WATCHING
                            excel.update_status(AppStatus.WATCHING, looking_for=looking_for)
                            log.info("WATCHING")
                except ExcelDisconnected as exc:
                    action = excel_failure_action(
                        exc, calculating=session.status == AppStatus.CALCULATING
                    )
                    if action == "error":
                        raise ExcelBridgeError(str(exc)) from exc
                    _enter_excel_wait(exc)
                    continue

            time.sleep(poll_sec)

    except StopRequested:
        _apply_stopped(looking_for)
        return 0
    except ExcelDisconnected as exc:
        log.error("ERROR | %s", exc)
        _write_error_cells(excel, cfg, str(exc), last_pnl=last_pnl)
        return 1
    except (ConfigError, ExcelBridgeError, SourceReaderError, SendError) as exc:
        log.error("ERROR | %s", exc)
        _write_error_cells(excel, cfg, str(exc), last_pnl=last_pnl)
        return 1
    except Exception as exc:
        log.exception("ERROR | unexpected: %s", exc)
        _write_error_cells(excel, cfg, str(exc), last_pnl=last_pnl)
        return 1
    finally:
        clear_pid_file(pid_path)
        if excel is not None:
            excel.close()


def main(argv: Optional[list[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)

    if args.serve:
        from app.web.server import run_local_web

        return run_local_web()
    if args.serve_admin:
        from admin.server import run_admin

        return run_admin()
    if args.run_profile:
        from app.watcher_profile import run_watcher_from_profile

        return run_watcher_from_profile()

    try:
        cfg = Config.load(args.config)
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        report_startup_error(args.config, str(exc))
        return 2

    setup_logger(cfg.log_path, level=cfg.log_level)

    try:
        if args.diagnose_source:
            return run_diagnose_source(cfg)
        if args.diagnose_send:
            return run_diagnose_send(cfg)
        if args.perf_summary:
            return run_perf_summary(cfg)
        if args.test_send:
            return run_test_send(cfg)
        if args.test_parser is not None:
            return run_test_parser(cfg, args.test_parser)
        return run_watcher(cfg)
    except (SourceReaderError, SendError, ExcelBridgeError, ConfigError) as exc:
        get_logger().error("ERROR | %s", exc)
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
