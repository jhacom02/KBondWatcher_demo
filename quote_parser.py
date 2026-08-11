"""Parse FORESTBOND chat lines into Quote objects."""

from __future__ import annotations

import re
from datetime import datetime
from typing import Optional

from models import Quote


class QuoteParseError(ValueError):
    """Raised when a line cannot be parsed as a target quote."""


_SENDER_TS = re.compile(
    r"^\s*(?P<sender>.+?)\s*\((?P<ts>\d{1,2}:\d{2}(?::\d{2})?)\)\s*[:：]\s*(?P<body>.*)$"
)

_TOKEN_PM = re.compile(r"(?P<digits>\d{2,3})(?P<side>[+-])")
_TOKEN_KR = re.compile(r"(?P<digits>\d{2,3})\s*(?P<side>사자|팔자)")


def build_target_pattern(target: str) -> re.Pattern[str]:
    escaped = re.escape(target)
    return re.compile(rf"(?<![\d-]){escaped}(?![\d-])")


def digits_to_yield(digits: str, yield_prefix: float) -> float:
    n = int(digits)
    if len(digits) == 2:
        return float(yield_prefix) + n / 100.0
    if len(digits) == 3:
        return float(yield_prefix) + n / 1000.0
    raise QuoteParseError(f"unsupported digit length: {digits!r}")


def parse_side_marker(marker: str) -> str:
    if marker in {"+", "사자"}:
        return "BUY"
    if marker in {"-", "팔자"}:
        return "SELL"
    raise QuoteParseError(f"unknown side marker: {marker!r}")


def _extract_meta(line: str) -> tuple[Optional[str], Optional[datetime], str]:
    match = _SENDER_TS.match(line)
    if not match:
        return None, None, line

    sender = match.group("sender").strip() or None
    ts_raw = match.group("ts")
    body = match.group("body")
    timestamp: Optional[datetime] = None
    try:
        if ts_raw.count(":") == 2:
            timestamp = datetime.strptime(ts_raw, "%H:%M:%S")
        else:
            timestamp = datetime.strptime(ts_raw, "%H:%M")
    except ValueError:
        timestamp = None
    return sender, timestamp, body


def find_quote_token(text_after_target: str) -> Optional[tuple[str, str, str]]:
    """
    Return (raw_token, digits, side_marker) from text following the instrument.
    """
    pm = _TOKEN_PM.search(text_after_target)
    kr = _TOKEN_KR.search(text_after_target)

    candidates: list[tuple[int, str, str, str]] = []
    if pm:
        candidates.append(
            (pm.start(), pm.group(0), pm.group("digits"), pm.group("side"))
        )
    if kr:
        candidates.append(
            (kr.start(), kr.group(0), kr.group("digits"), kr.group("side"))
        )
    if not candidates:
        return None
    candidates.sort(key=lambda item: item[0])
    _, raw_token, digits, side = candidates[0]
    return raw_token, digits, side


def parse_quote_line(
    line: str,
    target: str,
    yield_prefix: float,
    required_side: str = "ANY",
) -> Optional[Quote]:
    """
    Parse a single message line.

    Returns None if the line does not contain the target instrument / quote.
    """
    if not line or not line.strip():
        return None

    target_re = build_target_pattern(target)
    target_match = target_re.search(line)
    if not target_match:
        return None

    sender, timestamp, _ = _extract_meta(line)
    after = line[target_match.end() :]
    token_info = find_quote_token(after)
    if token_info is None:
        return None

    raw_token, digits, side_marker = token_info
    side = parse_side_marker(side_marker)
    required = (required_side or "ANY").upper()
    if required in {"BUY", "SELL"} and side != required:
        return None

    yield_value = digits_to_yield(digits, yield_prefix)
    return Quote(
        instrument=target,
        raw_line=line.strip(),
        raw_token=raw_token,
        yield_value=yield_value,
        side=side,
        timestamp=timestamp,
        sender=sender,
    )


def format_parser_result(quote: Quote) -> str:
    return (
        f"instrument = {quote.instrument}\n"
        f"yield = {quote.yield_value}\n"
        f"side = {quote.side}\n"
        f"raw_token = {quote.raw_token}"
    )
