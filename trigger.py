from __future__ import annotations

from models import Quote, TriggerResult


def evaluate(quote: Quote, pnl: float, threshold: float) -> TriggerResult:
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


def format_message(template: str, quote: Quote, pnl: float) -> str:
    return template.format(
        instrument=quote.instrument,
        raw_token=quote.raw_token,
        yield_value=quote.yield_value,
        side=quote.side,
        pnl=pnl,
        raw_line=quote.raw_line,
    )
