"""
FORESTBOND → Excel → K-Bond one-shot orchestration entrypoint.

Usage examples:
  python main.py --config config.env
  python main.py --config config.env --diagnose-kbond
  python main.py --config config.env --prefill-kbond
  python main.py --config config.env --diagnose-chrome
  python main.py --config config.env --test-parser "25-11 23+"
"""

from __future__ import annotations

import argparse
import sys
import time
from pathlib import Path
from typing import Optional

from config import Config, ConfigError
from excel_bridge import ExcelBridge, ExcelBridgeError
from forestbond_reader import ForestBondReader, ForestBondReaderError
from kbond_controller import (
    KBondError,
    format_diagnose_report,
    inspect_kbond,
    prefill_kbond,
)
from logger import get_logger, setup_logger
from models import AppStatus, Quote, TriggerResult, WatcherSession
from quote_parser import format_parser_result, parse_quote_line


class TriggerEngine:
    """POC: trigger when pnl >= threshold. Extensible for side/qty/age later."""

    def __init__(self, pnl_threshold: float) -> None:
        self.pnl_threshold = float(pnl_threshold)

    def evaluate(self, quote: Quote, pnl: float) -> TriggerResult:
        if pnl >= self.pnl_threshold:
            return TriggerResult(
                triggered=True,
                reason=f"pnl {pnl} >= threshold {self.pnl_threshold}",
                pnl=pnl,
                quote=quote,
            )
        return TriggerResult(
            triggered=False,
            reason=f"pnl {pnl} < threshold {self.pnl_threshold}",
            pnl=pnl,
            quote=quote,
        )


def clear_stop_flag(path: Path) -> None:
    try:
        if path.is_file():
            path.unlink()
    except OSError as exc:
        get_logger().warning("Failed to clear stop flag %s: %s", path, exc)


def stop_requested(path: Path) -> bool:
    return path.is_file()


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="FORESTBOND → Excel → K-Bond one-shot watcher"
    )
    parser.add_argument(
        "--config",
        default="config.env",
        help="Path to config.env (default: config.env)",
    )
    parser.add_argument(
        "--diagnose-chrome",
        action="store_true",
        help="Diagnose FORESTBOND Chrome UIA text reading and exit",
    )
    parser.add_argument(
        "--diagnose-kbond",
        action="store_true",
        help="Diagnose axis.exe HWND / click coordinate and exit",
    )
    parser.add_argument(
        "--prefill-kbond",
        action="store_true",
        help="Click K-Bond relative point and paste SEND_TEXT (no Enter)",
    )
    parser.add_argument(
        "--test-parser",
        metavar="LINE",
        help='Parse a sample quote line, e.g. "25-11 23+"',
    )
    return parser


def run_diagnose_chrome(cfg: Config) -> int:
    reader = ForestBondReader(chrome_title=cfg.chrome_title)
    print(reader.diagnose(max_messages=200))
    return 0


def run_diagnose_kbond(cfg: Config) -> int:
    from kbond_controller import resolve_kbond_pid, scan_kbond_candidates

    pid = resolve_kbond_pid(cfg.kbond_process_name, cfg.kbond_pid)
    print(scan_kbond_candidates(pid, cfg.kbond_process_name, cfg.kbond_window_title_contains))
    print("---- selected ----")
    info = inspect_kbond(
        process_name=cfg.kbond_process_name,
        configured_pid=cfg.kbond_pid,
        win_x=cfg.win_x,
        win_y=cfg.win_y,
        title_contains=cfg.kbond_window_title_contains,
    )
    print(format_diagnose_report(info, cfg.win_x, cfg.win_y))
    return 0


def run_prefill_kbond(cfg: Config) -> int:
    log = get_logger()
    info = prefill_kbond(
        process_name=cfg.kbond_process_name,
        configured_pid=cfg.kbond_pid,
        win_x=cfg.win_x,
        win_y=cfg.win_y,
        send_text=cfg.send_text,
        title_contains=cfg.kbond_window_title_contains,
    )
    log.info("READY_TO_SUBMIT")
    print(
        f"Prefill complete on HWND 0x{info.hwnd:08X}. "
        f'Pasted "{cfg.send_text}". Enter was NOT sent.'
    )
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
    trigger_engine = TriggerEngine(cfg.pnl_threshold)

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
        excel.update_status(AppStatus.WATCHING, last_action="watcher started")
        session.status = AppStatus.WATCHING
        log.info("WATCHING")

        reader = ForestBondReader(chrome_title=cfg.chrome_title)
        # Force window discovery early for clearer errors.
        reader.find_forestbond_window()
        reader.initialize_watermark(cfg.process_existing_on_start)

        poll_sec = cfg.poll_interval_ms / 1000.0

        while True:
            if stop_requested(cfg.stop_flag_path):
                session.status = AppStatus.STOPPED
                log.info("STOPPED")
                excel.update_status(AppStatus.STOPPED, last_action="stop flag detected")
                clear_stop_flag(cfg.stop_flag_path)
                return 0

            try:
                lines = reader.get_new_message_lines(
                    process_existing_on_start=cfg.process_existing_on_start
                )
            except ForestBondReaderError as exc:
                log.warning("FORESTBOND read error (retrying): %s", exc)
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

                fp = quote.fingerprint
                if fp in session.processed_fingerprints:
                    continue
                session.processed_fingerprints.add(fp)

                session.status = AppStatus.QUOTE_FOUND
                log.info(
                    "QUOTE_FOUND | %s | %s | %.3f | %s",
                    quote.instrument,
                    quote.raw_token,
                    quote.yield_value,
                    quote.side,
                )
                excel.update_status(
                    AppStatus.QUOTE_FOUND,
                    last_quote=f"{quote.instrument} {quote.raw_token} "
                    f"{quote.yield_value} {quote.side}",
                    last_action="quote found",
                )

                session.status = AppStatus.CALCULATING
                excel.update_status(AppStatus.CALCULATING, last_action="calculating")
                pnl = excel.write_yield_calculate_read_pnl(quote.yield_value)

                result = trigger_engine.evaluate(quote, pnl)
                if not result.triggered:
                    session.status = AppStatus.NO_TRIGGER
                    log.info("NO_TRIGGER | %s", result.reason)
                    excel.update_status(
                        AppStatus.NO_TRIGGER,
                        last_pnl=pnl,
                        last_action=result.reason,
                    )
                    session.status = AppStatus.WATCHING
                    excel.update_status(AppStatus.WATCHING, last_action="resume watching")
                    log.info("WATCHING")
                    continue

                session.status = AppStatus.TRIGGERED
                log.info("TRIGGERED")
                excel.update_status(
                    AppStatus.TRIGGERED,
                    last_pnl=pnl,
                    last_action=result.reason,
                )

                session.status = AppStatus.PREFILLING
                excel.update_status(AppStatus.PREFILLING, last_action="prefill kbond")
                prefill_kbond(
                    process_name=cfg.kbond_process_name,
                    configured_pid=cfg.kbond_pid,
                    win_x=cfg.win_x,
                    win_y=cfg.win_y,
                    send_text=cfg.send_text,
                    title_contains=cfg.kbond_window_title_contains,
                )

                session.status = AppStatus.READY_TO_SUBMIT
                log.info("READY_TO_SUBMIT")
                excel.update_status(
                    AppStatus.READY_TO_SUBMIT,
                    last_pnl=pnl,
                    last_action='prefilled SEND_TEXT; press Enter manually',
                )
                log.info("EXIT")
                excel.update_status(AppStatus.DONE, last_action="watcher exit")
                return 0

            time.sleep(poll_sec)

    except (ConfigError, ExcelBridgeError, ForestBondReaderError, KBondError) as exc:
        log.error("ERROR | %s", exc)
        if excel is not None:
            excel.update_status(AppStatus.ERROR, last_action=str(exc)[:200])
        return 1
    except Exception as exc:  # noqa: BLE001
        log.exception("ERROR | unexpected: %s", exc)
        if excel is not None:
            excel.update_status(AppStatus.ERROR, last_action=str(exc)[:200])
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
        if args.diagnose_chrome:
            return run_diagnose_chrome(cfg)
        if args.diagnose_kbond:
            return run_diagnose_kbond(cfg)
        if args.prefill_kbond:
            return run_prefill_kbond(cfg)
        if args.test_parser is not None:
            return run_test_parser(cfg, args.test_parser)
        return run_watcher(cfg)
    except (ForestBondReaderError, KBondError, ExcelBridgeError, ConfigError) as exc:
        get_logger().error("ERROR | %s", exc)
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    sys.exit(main())
