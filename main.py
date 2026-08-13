from __future__ import annotations

import argparse
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Optional

from config import Config, ConfigError
from excel import ExcelBridge, ExcelBridgeError, InstrumentSlot
from source import SourceReaderError, create_source_reader, format_parser_result, parse_quote_line
from core import (
    AppStatus,
    Quote,
    WatcherSession,
    evaluate,
    format_message,
    get_logger,
    setup_logger,
)
import send
from send import SendError


def clear_stop_flag(path: Path) -> None:
    try:
        if path.is_file():
            path.unlink()
    except OSError as exc:
        get_logger().warning("Failed to clear stop flag %s: %s", path, exc)


def stop_requested(path: Path) -> bool:
    return path.is_file()


def format_last_action(action: str) -> str:
    stamp = datetime.now().strftime("%H:%M:%S")
    return f"({stamp}) {action}"


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=".env")
    parser.add_argument("--diagnose-source", action="store_true")
    parser.add_argument("--diagnose-send", action="store_true")
    parser.add_argument("--test-send", action="store_true")
    parser.add_argument("--test-parser", metavar="LINE")
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
        slots, looking_for = excel.load_slots()
    finally:
        excel.close()
    for slot in slots:
        quote = parse_quote_line(
            line=line,
            target=slot.instrument,
            yield_prefix=slot.yield_prefix,
            required_side=slot.required_side,
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
        )
        if quote is not None:
            return quote, slot
    return None, None


def run_watcher(cfg: Config) -> int:
    log = get_logger()
    session = WatcherSession(status=AppStatus.STARTING)
    excel: Optional[ExcelBridge] = None
    clear_stop_flag(cfg.stop_flag_path)

    try:
        excel = _build_excel(cfg)
        excel.connect()
        slots, looking_for = excel.load_slots()
        excel.update_status(
            AppStatus.WATCHING,
            looking_for=looking_for,
            last_action=format_last_action("Start Successful"),
        )
        session.status = AppStatus.WATCHING
        log.info(
            "WATCHING | mode=%s source=%s/%s send=%s/%s looking_for=%s slots=%s",
            cfg.mode,
            cfg.source_process_name or "(uia)",
            cfg.source_window_title,
            cfg.send_process_name,
            cfg.send_window_title,
            looking_for,
            [(s.instrument, s.row, s.required_side) for s in slots],
        )

        reader = create_source_reader(cfg)
        reader.find_source_window()
        send.ensure_target_window(cfg)
        reader.initialize_watermark(cfg.process_existing_on_start)
        poll_sec = cfg.poll_interval_ms / 1000.0

        while True:
            if stop_requested(cfg.stop_flag_path):
                session.status = AppStatus.STOPPED
                log.info("STOPPED")
                excel.update_status(
                    AppStatus.STOPPED,
                    looking_for=looking_for,
                    last_action=format_last_action("Stopped"),
                )
                clear_stop_flag(cfg.stop_flag_path)
                return 0

            lines = reader.get_new_message_lines(
                process_existing_on_start=cfg.process_existing_on_start
            )

            for line in lines:
                quote, slot = _match_quote(line, slots)
                if quote is None or slot is None:
                    continue
                if quote.fingerprint in session.processed_fingerprints:
                    continue
                session.processed_fingerprints.add(quote.fingerprint)

                session.status = AppStatus.QUOTE_FOUND
                log.info(
                    "QUOTE_FOUND | %s | raw_token=%s | %.3f | %s | row=%s | raw_line=%s",
                    quote.instrument,
                    quote.raw_token,
                    quote.yield_value,
                    quote.side,
                    slot.row,
                    quote.raw_line,
                )

                session.status = AppStatus.CALCULATING
                pnl = excel.write_yield_read_pnl(
                    slot.input_cell,
                    slot.pnl_cell,
                    quote.yield_value,
                )
                result = evaluate(
                    quote,
                    pnl,
                    cfg.pnl_threshold,
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
                    continue

                session.status = AppStatus.TRIGGERED
                log.info("TRIGGERED")
                text = format_message(cfg.message_template, quote, pnl)
                session.status = AppStatus.SENDING
                send.send_text(text, cfg)

                session.status = AppStatus.SENT
                log.info("SENT")
                excel.update_status(
                    AppStatus.SENT,
                    looking_for=looking_for,
                    last_quote=f"{quote.instrument} {quote.raw_token}",
                    last_pnl=pnl,
                    last_action=format_last_action(f"Message Sent: {text}"),
                )
                # <<<< 2 lines are for one-shot >>>>
                # log.info("EXIT")
                # return 0
                # <<<< test-loop: reseed then keep watching >>>>
                reader.reseed_watermark_from_visible()
                session.status = AppStatus.WATCHING
                excel.update_status(AppStatus.WATCHING, looking_for=looking_for)
                log.info("WATCHING")
                continue

            time.sleep(poll_sec)

    except (ConfigError, ExcelBridgeError, SourceReaderError, SendError) as exc:
        log.error("ERROR | %s", exc)
        if excel is not None:
            excel.update_status(
                AppStatus.ERROR,
                last_action=format_last_action(f"Python Error: {str(exc)[:200]}"),
            )
        return 1
    except Exception as exc:
        log.exception("ERROR | unexpected: %s", exc)
        if excel is not None:
            excel.update_status(
                AppStatus.ERROR,
                last_action=format_last_action(f"Python Error: {str(exc)[:200]}"),
            )
        return 1
    finally:
        if excel is not None:
            excel.close()


def main(argv: Optional[list[str]] = None) -> int:
    args = build_arg_parser().parse_args(argv)
    try:
        cfg = Config.load(args.config)
    except ConfigError as exc:
        print(f"Config error: {exc}", file=sys.stderr)
        return 2

    setup_logger(cfg.log_path, level=cfg.log_level)

    try:
        if args.diagnose_source:
            return run_diagnose_source(cfg)
        if args.diagnose_send:
            return run_diagnose_send(cfg)
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
