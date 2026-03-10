from __future__ import annotations

from pathlib import Path

import openpyxl
import pandas as pd


def build_stock_master_from_excel(source_path: str, output_path: str) -> pd.DataFrame:
    source = Path(source_path)
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)

    wb = openpyxl.load_workbook(source, read_only=True, data_only=True)
    ws = wb["종목"]

    rows = []
    for name, code in ws.iter_rows(min_row=2, values_only=True):
        if not name or not code:
            continue
        symbol = str(code).strip().zfill(6)
        rows.append(
            {
                "symbol": symbol,
                "name": str(name).strip(),
                "market": "KR",
                "enabled": 1,
                "source_file": str(source),
            }
        )

    df = pd.DataFrame(rows).drop_duplicates(subset=["symbol"]).sort_values("symbol")
    df.to_csv(output, index=False, encoding="utf-8-sig")
    return df


def load_stock_master(path: str) -> pd.DataFrame:
    df = pd.read_csv(path, dtype={"symbol": str})
    df["symbol"] = df["symbol"].str.zfill(6)
    if "enabled" in df.columns:
        df = df[df["enabled"].fillna(1).astype(int) == 1]
    return df.reset_index(drop=True)
