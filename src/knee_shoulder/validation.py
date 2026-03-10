from __future__ import annotations

from pathlib import Path

import pandas as pd

from .storage import load_existing_history


def _forward_return(prices: pd.Series, entry_index: int, forward_days: int) -> float | None:
    target = entry_index + forward_days
    if target >= len(prices):
        return None
    entry = prices.iloc[entry_index]
    future = prices.iloc[target]
    if not entry:
        return None
    return round(((future / entry) - 1.0) * 100.0, 2)


def build_validation_rows(signals: pd.DataFrame, raw_dir: str, forward_days: list[int]) -> pd.DataFrame:
    rows = []
    for signal in signals.itertuples(index=False):
        history_path = Path(raw_dir) / f"{signal.symbol}.csv"
        history = load_existing_history(history_path)
        if history.empty:
            continue
        history = history.sort_values("date").reset_index(drop=True)
        matches = history.index[history["date"] == signal.date].tolist()
        if not matches:
            continue
        idx = matches[0]
        close_series = history["close"].astype(float)
        row = {
            "signal_date": signal.date,
            "symbol": signal.symbol,
            "name": signal.name,
            "knee_score": signal.knee_score,
            "shoulder_score": signal.shoulder_score,
        }
        future_returns = {}
        for days in forward_days:
            future_returns[f"ret_{days}d"] = _forward_return(close_series, idx, days)
        row.update(future_returns)
        row["knee_success"] = int((future_returns.get("ret_5d") or -999) >= 3.0)
        row["shoulder_success"] = int((future_returns.get("ret_5d") or 999) <= -3.0)
        rows.append(row)
    return pd.DataFrame(rows)
