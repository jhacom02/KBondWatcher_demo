from __future__ import annotations

from models import Quote, TriggerResult

LOOKING_OFFER = "OFFER"
LOOKING_BID = "BID"


def looking_for_from_qty(qty: float) -> tuple[str, str]:
    if abs(qty - (-100.0)) < 1e-9:
        return LOOKING_BID, "BUY"
    if abs(qty - 100.0) < 1e-9:
        return LOOKING_OFFER, "SELL"
    raise ValueError(f"qty must be -100 or +100, got {qty!r}")


def flip_side_token(raw_token: str) -> str:
    token = (raw_token or "").strip()
    if token.endswith("+"):
        return token[:-1] + "-"
    if token.endswith("-"):
        return token[:-1] + "+"
    if token.endswith("사자"):
        return token[: -len("사자")] + "팔자"
    if token.endswith("팔자"):
        return token[: -len("팔자")] + "사자"
    raise ValueError(f"cannot flip side token: {raw_token!r}")


def evaluate(
    quote: Quote,
    pnl: float,
    threshold: float,
    looking_for: str,
) -> TriggerResult:
    looking = (looking_for or "").upper()
    if looking == LOOKING_BID:
        if pnl >= threshold:
            return TriggerResult(
                triggered=True,
                reason=f"pnl {pnl} >= threshold {threshold}",
                pnl=pnl,
                quote=quote,
            )
        return TriggerResult(
            triggered=False,
            reason=f"pnl {pnl} < threshold {threshold}",
            pnl=pnl,
            quote=quote,
        )
    if looking == LOOKING_OFFER:
        limit = -abs(threshold)
        if pnl <= limit:
            return TriggerResult(
                triggered=True,
                reason=f"pnl {pnl} <= {limit}",
                pnl=pnl,
                quote=quote,
            )
        return TriggerResult(
            triggered=False,
            reason=f"pnl {pnl} > {limit}",
            pnl=pnl,
            quote=quote,
        )
    raise ValueError(f"looking_for must be BID or OFFER, got {looking_for!r}")


def format_message(template: str, quote: Quote, pnl: float) -> str:
    confirm_token = flip_side_token(quote.raw_token)
    return template.format(
        instrument=quote.instrument,
        raw_token=quote.raw_token,
        confirm_token=confirm_token,
        yield_value=quote.yield_value,
        side=quote.side,
        pnl=pnl,
        raw_line=quote.raw_line,
    )
