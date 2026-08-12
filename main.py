from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path
from typing import Optional

from config import Config, ConfigError
from excel_bridge import ExcelBridge, ExcelBridgeError
from source_reader import SourceReader, SourceReaderError
from logger import get_logger, setup_logger
import send_ui
from send_ui import SendError
from models import AppStatus, Quote, WatcherSession
from quote_parser import format_parser_result, parse_quote_line
from trigger import evaluate, format_message


def clear_stop_flag(path: Path) -> None:
    try:
        if path.is_file():
            path.unlink()
    except OSError as exc:
        get_logger().warning("Failed to clear stop flag %s: %s", path, exc)


def stop_requested(path: Path) -> bool:
    return path.is_file()


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", default=".env")
    parser.add_argument("--diagnose-source", action="store_true")
    parser.add_argument("--diagnose-send", action="store_true")
    parser.add_argument("--test-send", action="store_true")
    parser.add_argument("--test-parser", metavar="LINE")
    return parser


def run_diagnose_source(cfg: Config) -> int:
    reader = SourceReader(
        source_window_title=cfg.source_window_title,
        source_process_name=cfg.source_process_name,
    )
    print(reader.diagnose(max_messages=200))
    return 0


def run_diagnose_send(cfg: Config) -> int:
    print(send_ui.diagnose(cfg))
    return 0


def run_test_send(cfg: Config) -> int:
    sample = Quote(
        instrument=cfg.target,
        raw_line=f"{cfg.target} test",
        raw_token="00+",
        yield_value=float(cfg.yield_prefix),
        side="BUY",
    )
    text = format_message(cfg.message_template, sample, pnl=0.0)
    send_ui.send_text(text, cfg)
    print(f"sent to {cfg.send_window_title!r}: {text}")
    return 0


def run_test_parser(cfg: Config, line: str) -> int:
    quote = parse_quote_line(
        line=line,
        target=cfg.target,
        yield_prefix=cfg.yield_prefix,
        required_side=cfg.required_side,
    )
    if quote is None:
        print(f"No match for TARGET={cfg.target!r} in line: {line!r}")
        return 1
    print(format_parser_result(quote))
    return 0


def run_watcher(cfg: Config) -> int:
    log = get_logger()
    session = WatcherSession(status=AppStatus.STARTING)
    excel: Optional[ExcelBridge] = None
    clear_stop_flag(cfg.stop_flag_path)

    try:
        excel = ExcelBridge(
            workbook_name=cfg.excel_workbook,
            sheet_name=cfg.excel_sheet,
            input_cell=cfg.excel_input_cell,
            pnl_cell=cfg.excel_pnl_cell,
            status_cell=cfg.excel_status_cell,
            last_quote_cell=cfg.excel_last_quote_cell,
            last_pnl_cell=cfg.excel_last_pnl_cell,
            last_action_cell=cfg.excel_last_action_cell,
        )
        excel.connect()
        excel.update_status(AppStatus.WATCHING, last_action="Start Successful")
        session.status = AppStatus.WATCHING
        log.info("WATCHING")

        reader = SourceReader(
            source_window_title=cfg.source_window_title,
            source_process_name=cfg.source_process_name,
        )
        reader.find_source_window()
        reader.initialize_watermark(cfg.process_existing_on_start)
        poll_sec = cfg.poll_interval_ms / 1000.0

        while True:
            if stop_requested(cfg.stop_flag_path):
                session.status = AppStatus.STOPPED
                log.info("STOPPED")
                excel.update_status(AppStatus.STOPPED, last_action="Stopped")
                clear_stop_flag(cfg.stop_flag_path)
                return 0

            try:
                lines = reader.get_new_message_lines(
                    process_existing_on_start=cfg.process_existing_on_start
                )
            except SourceReaderError as exc:
                log.warning("source read error: %s", exc)
                time.sleep(poll_sec)
                continue

            for line in lines:
                quote = parse_quote_line(
                    line=line,
                    target=cfg.target,
                    yield_prefix=cfg.yield_prefix,
                    required_side=cfg.required_side,
                )
                if quote is None:
                    continue
                if quote.fingerprint in session.processed_fingerprints:
                    continue
                session.processed_fingerprints.add(quote.fingerprint)

                session.status = AppStatus.QUOTE_FOUND
                log.info(
                    "QUOTE_FOUND | %s | %s | %.3f | %s",
                    quote.instrument,
                    quote.raw_token,
                    quote.yield_value,
                    quote.side,
                )

                session.status = AppStatus.CALCULATING
                pnl = excel.write_yield_read_pnl(quote.yield_value)
                result = evaluate(quote, pnl, cfg.pnl_threshold)
                pnl_text = f"{int(round(pnl)):,}"

                if not result.triggered:
                    session.status = AppStatus.NO_TRIGGER
                    log.info("NO_TRIGGER | %s", result.reason)
                    session.status = AppStatus.WATCHING
                    excel.update_status(
                        AppStatus.WATCHING,
                        last_quote=f"{quote.instrument} {quote.raw_token}",
                        last_pnl=pnl,
                        last_action=(
                            f"Quote Passed: {quote.instrument} {quote.raw_token} "
                            f"(pnl={pnl_text})"
                        ),
                    )
                    log.info("WATCHING")
                    continue

                session.status = AppStatus.TRIGGERED
                log.info("TRIGGERED")
                text = format_message(cfg.message_template, quote, pnl)
                session.status = AppStatus.SENDING
                send_ui.send_text(text, cfg)

                session.status = AppStatus.SENT
                log.info("SENT")
                excel.update_status(
                    AppStatus.SENT,
                    last_quote=f"{quote.instrument} {quote.raw_token}",
                    last_pnl=pnl,
                    last_action=f"Message Sent: {text}",
                )
                log.info("EXIT")
                return 0

            time.sleep(poll_sec)

    except (ConfigError, ExcelBridgeError, SourceReaderError, SendError) as exc:
        log.error("ERROR | %s", exc)
        if excel is not None:
            excel.update_status(
                AppStatus.ERROR,
                last_action=f"Python Error: {str(exc)[:200]}",
            )
        return 1
    except Exception as exc:
        log.exception("ERROR | unexpected: %s", exc)
        if excel is not None:
            excel.update_status(
                AppStatus.ERROR,
                last_action=f"Python Error: {str(exc)[:200]}",
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
