from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path
from typing import Optional

from excel import (
    ExcelBridgeError,
    InstrumentSlot,
    StopRequested,
)
from source import SourceLine, SourceReaderError, message_fingerprint, parse_quote_line
from core import Quote, get_logger


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


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--run-profile", action="store_true")
    parser.add_argument("--serve", action="store_true")
    parser.add_argument("--serve-admin", action="store_true")
    return parser


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
    return "error"


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

    build_arg_parser().print_help()
    return 2


if __name__ == "__main__":
    sys.exit(main())
