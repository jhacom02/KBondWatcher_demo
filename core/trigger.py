from __future__ import annotations

from .models import Quote, TriggerResult

LOOKING_OFFER = "OFFER"
LOOKING_BID = "BID"


def qty_magnitude(qty: float) -> int:
    value = float(qty)
    if abs(value) < 1e-9:
        raise ValueError(f"qty must be a non-zero integer, got {qty!r}")
    rounded = round(abs(value))
    if abs(abs(value) - rounded) > 1e-6:
        raise ValueError(f"qty must be a non-zero integer, got {qty!r}")
    return int(rounded)


def looking_for_from_qty(qty: float) -> tuple[str, str]:
    qty_magnitude(qty)
    if float(qty) < 0:
        return LOOKING_BID, "BUY"
    return LOOKING_OFFER, "SELL"


def format_looking_for_label(instrument: str, looking_for: str) -> str:
    return f"{instrument} / {looking_for}"


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


def pnl_outside_sanity_band(pnl: float, threshold: float, band: float) -> bool:
    return abs(float(pnl) - float(threshold)) > float(band)


def evaluate(
    quote: Quote,
    pnl: float,
    threshold: float,
    threshold_op: str = "<=",
) -> TriggerResult:
    op = (threshold_op or "<=").strip()
    if op == "<=":
        if pnl <= threshold:
            return TriggerResult(
                triggered=True,
                reason=f"pnl {pnl} <= threshold {threshold}",
                pnl=pnl,
                quote=quote,
            )
        return TriggerResult(
            triggered=False,
            reason=f"pnl {pnl} > threshold {threshold}",
            pnl=pnl,
            quote=quote,
        )
    if op == ">=":
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
    raise ValueError(f"threshold_op must be <= or >=, got {threshold_op!r}")


def format_message(template: str, quote: Quote, pnl: float) -> str:
    confirm_token = flip_side_token(quote.raw_token)
    if quote.quantity != 100:
        confirm_token = f"{confirm_token} {quote.quantity}억"
    return template.format(
        instrument=quote.instrument,
        raw_token=quote.raw_token,
        confirm_token=confirm_token,
        yield_value=quote.yield_value,
        side=quote.side,
        pnl=pnl,
        raw_line=quote.raw_line,
        quantity=quote.quantity,
    )
