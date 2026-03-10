from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.knee_shoulder.config import load_config
from src.knee_shoulder.storage import load_existing_history, load_validation_history


st.set_page_config(page_title="Knee Shoulder Monitor", layout="wide")
st.title("Knee/Shoulder Stock Monitor")
st.caption("Daily close-based reversal monitoring dashboard for Korean stocks.")

config = load_config()
paths = config["paths"]


def render_candidate_help(title: str, score_label: str, reasons_label: str) -> None:
    with st.popover("!"):
        st.markdown(f"**{title} columns**")
        st.markdown("- `symbol`: 종목코드")
        st.markdown("- `name`: 종목명")
        st.markdown("- `close`: 분석 기준일 종가")
        st.markdown("- `pct_change`: 전일 대비 등락률(%)")
        st.markdown(f"- `{score_label}`: 반전 후보 점수(0~100)")
        st.markdown(
            "- `knee_grade` / `shoulder_grade`: `Strong`, `Watch`, `Neutral` 등급"
        )
        st.markdown("- `vol_ratio_20`: 당일 거래량 / 20일 평균 거래량")
        st.markdown(f"- `{reasons_label}`: 점수에 반영된 핵심 이유 요약")


def load_latest_signals(signal_dir: str) -> tuple[pd.DataFrame, str | None]:
    files = sorted(Path(signal_dir).glob("*_signals.csv"))
    if not files:
        return pd.DataFrame(), None
    latest = files[-1]
    return pd.read_csv(latest, dtype={"symbol": str}), latest.stem.replace("_signals", "")


def prepare_history_for_chart(history: pd.DataFrame) -> pd.DataFrame:
    frame = history.copy()
    frame["date"] = pd.to_datetime(frame["date"].astype(str), format="%Y%m%d", errors="coerce")
    frame = frame.dropna(subset=["date"]).sort_values("date").reset_index(drop=True)
    return frame


def format_candidate_view(frame: pd.DataFrame, score_column: str, reasons_column: str) -> pd.DataFrame:
    view = frame[
        ["symbol", "name", "close", "pct_change", score_column, f"{score_column.split('_')[0]}_grade", "vol_ratio_20", reasons_column]
    ].copy()
    view["close"] = pd.to_numeric(view["close"], errors="coerce").map(lambda v: f"{int(v):,}" if pd.notna(v) else "")
    return view


signals_df, signal_date = load_latest_signals(paths["signal_dir"])
validation_df = load_validation_history(Path(paths["validation_file"]))

if signals_df.empty:
    st.warning("No signal file found yet. Run `python3 run_daily.py` first.")
    st.stop()

header_cols = st.columns(4)
analysis_date = signal_date or "-"
run_at = signals_df["run_at"].iloc[0] if "run_at" in signals_df.columns and not signals_df.empty else "-"
header_cols[0].metric("Analysis Date", analysis_date)
header_cols[1].metric("Knee Strong", int((signals_df["knee_grade"] == "Strong").sum()))
header_cols[2].metric("Shoulder Strong", int((signals_df["shoulder_grade"] == "Strong").sum()))
header_cols[3].metric("Run At", run_at)

knee_view = signals_df[signals_df["knee_score"] >= config["runtime"]["signal_threshold"]].copy()
shoulder_view = signals_df[signals_df["shoulder_score"] >= config["runtime"]["signal_threshold"]].copy()

knee_header_col, knee_help_col = st.columns([20, 1])
with knee_header_col:
    st.subheader("Knee Candidates")
with knee_help_col:
    render_candidate_help("Knee Candidates", "knee_score", "knee_reasons")
st.dataframe(
    format_candidate_view(knee_view, "knee_score", "knee_reasons"),
    use_container_width=True,
    hide_index=True,
)

shoulder_header_col, shoulder_help_col = st.columns([20, 1])
with shoulder_header_col:
    st.subheader("Shoulder Candidates")
with shoulder_help_col:
    render_candidate_help("Shoulder Candidates", "shoulder_score", "shoulder_reasons")
st.dataframe(
    format_candidate_view(shoulder_view, "shoulder_score", "shoulder_reasons"),
    use_container_width=True,
    hide_index=True,
)

st.subheader("Symbol Detail")
st.caption("후보 종목만 선택할 수 있습니다. Knee / Shoulder를 나눠서 고르세요.")

knee_options = (knee_view["symbol"] + " | " + knee_view["name"]).tolist()
shoulder_options = (shoulder_view["symbol"] + " | " + shoulder_view["name"]).tolist()

selector_col1, selector_col2 = st.columns(2)
with selector_col1:
    knee_selected = st.radio(
        "Knee Candidate",
        options=knee_options if knee_options else ["후보 없음"],
        key="knee_candidate_radio",
    )
with selector_col2:
    shoulder_selected = st.radio(
        "Shoulder Candidate",
        options=shoulder_options if shoulder_options else ["후보 없음"],
        key="shoulder_candidate_radio",
    )

selected_option = None
if knee_selected != "후보 없음":
    selected_option = knee_selected
elif shoulder_selected != "후보 없음":
    selected_option = shoulder_selected
else:
    st.info("현재 상세 보기로 선택할 후보 종목이 없습니다.")
    st.stop()

selected_symbol = selected_option.split(" | ", 1)[0]
selected_row = signals_df[signals_df["symbol"] == selected_symbol].iloc[0]
history = load_existing_history(Path(paths["raw_dir"]) / f"{selected_symbol}.csv")

if not history.empty:
    history = prepare_history_for_chart(history)
    figure = go.Figure()
    figure.add_trace(go.Scatter(x=history["date"], y=history["close"], mode="lines", name="Close"))
    if "ma_20" in history.columns:
        figure.add_trace(go.Scatter(x=history["date"], y=history["ma_20"], mode="lines", name="MA20"))
    figure.update_layout(
        height=420,
        margin=dict(l=20, r=20, t=20, b=20),
        xaxis=dict(
            title="Date",
            tickformat="%Y-%m-%d",
            type="date",
        ),
        yaxis=dict(
            title="Close",
            tickformat=",d",
        ),
    )
    st.plotly_chart(figure, use_container_width=True)

detail_cols = st.columns(2)
detail_cols[0].write(
    {
        "close": f"{int(selected_row['close']):,}",
        "knee_score": int(selected_row["knee_score"]),
        "knee_grade": selected_row["knee_grade"],
        "knee_reasons": selected_row["knee_reasons"],
        "knee_confirmed": bool(selected_row["knee_confirmed"]),
    }
)
detail_cols[1].write(
    {
        "close": f"{int(selected_row['close']):,}",
        "shoulder_score": int(selected_row["shoulder_score"]),
        "shoulder_grade": selected_row["shoulder_grade"],
        "shoulder_reasons": selected_row["shoulder_reasons"],
        "shoulder_confirmed": bool(selected_row["shoulder_confirmed"]),
    }
)

st.subheader("Validation")
if not validation_df.empty:
    symbol_validation = validation_df[validation_df["symbol"] == selected_symbol]
    st.dataframe(symbol_validation, use_container_width=True, hide_index=True)
else:
    st.info("Validation data will appear after enough forward days have accumulated.")
