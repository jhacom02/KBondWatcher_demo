from __future__ import annotations

import re
from datetime import datetime
from typing import Optional

from core.models import Quote


class QuoteParseError(ValueError):
    pass


_SENDER_TS = re.compile(
    r"^\s*(?P<sender>.+?)\s*\((?P<ts>\d{1,2}:\d{2}(?::\d{2})?)\)\s*[:：]\s*(?P<body>.*)$"
)
_PRICE_SIDE = re.compile(
    r"^\s+(?P<price>\d{2,3})\s*(?P<side>[+-]|사자|팔자)\s*"
)


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


def parse_quote_line(
    line: str,
    target: str,
    yield_prefix: float,
    required_side: str = "ANY",
) -> Optional[Quote]:
    if not line or not line.strip():
        return None
    target_re = build_target_pattern(target)
    target_match = target_re.search(line)
    if not target_match:
        return None
    sender, timestamp, _ = _extract_meta(line)
    after = line[target_match.end() :]
    token_match = _PRICE_SIDE.match(after)
    if token_match is None:
        return None
    rest = after[token_match.end() :]
    rest_stripped = rest.strip()
    if rest_stripped and not rest_stripped.startswith(("(", "*")):
        return None
    price = token_match.group("price")
    side_marker = token_match.group("side")
    side = parse_side_marker(side_marker)
    required = (required_side or "ANY").upper()
    if required in {"BUY", "SELL"} and side != required:
        return None
    yield_value = digits_to_yield(price, yield_prefix)
    return Quote(
        instrument=target,
        raw_line=line.strip(),
        raw_token=token_match.group(0).strip(),
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
