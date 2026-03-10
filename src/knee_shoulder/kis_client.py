from __future__ import annotations

import time
from dataclasses import dataclass

import pandas as pd
import requests


@dataclass
class KisAuth:
    app_key: str
    app_secret: str
    base_url: str


def issue_access_token(auth: KisAuth) -> str:
    url = f"{auth.base_url}/oauth2/tokenP"
    payload = {
        "grant_type": "client_credentials",
        "appkey": auth.app_key,
        "appsecret": auth.app_secret,
    }
    headers = {
        "content-type": "application/json",
        "appKey": auth.app_key,
        "appSecret": auth.app_secret,
    }
    response = requests.post(url, headers=headers, json=payload, timeout=15)
    response.raise_for_status()
    data = response.json()
    token = data.get("access_token")
    if not token:
        raise ValueError("KIS token response did not include access_token")
    return token


def _base_headers(auth: KisAuth, access_token: str, tr_id: str) -> dict:
    return {
        "content-type": "application/json; charset=utf-8",
        "authorization": f"Bearer {access_token}",
        "appkey": auth.app_key,
        "appsecret": auth.app_secret,
        "tr_id": tr_id,
        "custtype": "P",
    }


def fetch_daily_history(
    auth: KisAuth,
    access_token: str,
    symbol: str,
    start_date: str,
    end_date: str,
) -> pd.DataFrame:
    url = f"{auth.base_url}/uapi/domestic-stock/v1/quotations/inquire-daily-itemchartprice"
    params = {
        "FID_COND_MRKT_DIV_CODE": "J",
        "FID_INPUT_ISCD": symbol,
        "FID_PERIOD_DIV_CODE": "D",
        "FID_ORG_ADJ_PRC": "1",
        "FID_INPUT_DATE_1": start_date,
        "FID_INPUT_DATE_2": end_date,
        "FID_COMP_ICD": symbol,
    }
    response = requests.get(
        url,
        headers=_base_headers(auth, access_token, "FHKST03010100"),
        params=params,
        timeout=20,
    )
    response.raise_for_status()
    data = response.json()
    rows = data.get("output2") or []
    records = []
    for item in rows:
        records.append(
            {
                "date": item.get("stck_bsop_date", ""),
                "open": int(item.get("stck_oprc", "0") or 0),
                "high": int(item.get("stck_hgpr", "0") or 0),
                "low": int(item.get("stck_lwpr", "0") or 0),
                "close": int(item.get("stck_clpr", "0") or 0),
                "volume": int(item.get("acml_vol", "0") or 0),
                "turnover": int(item.get("acml_tr_pbmn", "0") or 0),
            }
        )
    frame = pd.DataFrame.from_records(records)
    if frame.empty:
        return frame
    frame = frame.sort_values("date").drop_duplicates(subset=["date"]).reset_index(drop=True)
    return frame


def fetch_investor_trade_by_stock_daily(
    auth: KisAuth,
    access_token: str,
    symbol: str,
    start_date: str,
    end_date: str,
) -> pd.DataFrame:
    url = f"{auth.base_url}/uapi/domestic-stock/v1/quotations/investor-trade-by-stock-daily"
    params = {
        "FID_COND_MRKT_DIV_CODE": "J",
        "FID_INPUT_ISCD": symbol,
        "FID_INPUT_DATE_1": start_date,
        "FID_INPUT_DATE_2": end_date,
        "FID_PERIOD_DIV_CODE": "D",
    }
    response = requests.get(
        url,
        headers=_base_headers(auth, access_token, "FHKST66300000"),
        params=params,
        timeout=20,
    )
    response.raise_for_status()
    data = response.json()
    rows = data.get("output") or data.get("output1") or []
    if not rows:
        return pd.DataFrame()
    frame = pd.DataFrame(rows)
    return frame


def throttle(seconds: float) -> None:
    if seconds > 0:
        time.sleep(seconds)
