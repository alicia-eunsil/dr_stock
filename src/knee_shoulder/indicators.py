from __future__ import annotations

import pandas as pd


def add_indicators(df: pd.DataFrame) -> pd.DataFrame:
    frame = df.copy()
    close = frame["close"].astype(float)
    volume = frame["volume"].astype(float)

    for window in (5, 20, 60, 120):
        frame[f"ma_{window}"] = close.rolling(window).mean()

    frame["vol_ma_20"] = volume.rolling(20).mean()
    frame["vol_ratio_20"] = volume / frame["vol_ma_20"]
    frame["low_20"] = close.rolling(20).min()
    frame["high_20"] = close.rolling(20).max()
    frame["low_60"] = close.rolling(60).min()
    frame["high_60"] = close.rolling(60).max()
    frame["dist_from_low_20_pct"] = ((close / frame["low_20"]) - 1.0) * 100.0
    frame["dist_from_high_20_pct"] = ((close / frame["high_20"]) - 1.0) * 100.0

    delta = close.diff()
    gain = delta.clip(lower=0)
    loss = -delta.clip(upper=0)
    avg_gain = gain.ewm(alpha=1 / 14, adjust=False, min_periods=14).mean()
    avg_loss = loss.ewm(alpha=1 / 14, adjust=False, min_periods=14).mean()
    rs = avg_gain / avg_loss.replace(0, pd.NA)
    frame["rsi_14"] = 100 - (100 / (1 + rs))

    ema_12 = close.ewm(span=12, adjust=False).mean()
    ema_26 = close.ewm(span=26, adjust=False).mean()
    frame["macd"] = ema_12 - ema_26
    frame["macd_signal"] = frame["macd"].ewm(span=9, adjust=False).mean()
    frame["macd_hist"] = frame["macd"] - frame["macd_signal"]
    return frame
