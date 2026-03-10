from __future__ import annotations

from pathlib import Path

import pandas as pd
import plotly.graph_objects as go
import streamlit as st

from src.knee_shoulder.config import load_config
from src.knee_shoulder.storage import load_existing_history, load_validation_history


st.set_page_config(page_title="Knee Shoulder Monitor", layout="wide")
st.markdown(
    """
    <style>
    :root {
        --cool-bg: #eef5fb;
        --cool-surface: rgba(255, 255, 255, 0.88);
        --cool-border: #bfd2e6;
        --cool-ink: #12324a;
        --cool-muted: #55758f;
        --cool-accent: #2f6fa3;
        --cool-accent-soft: #dbeaf7;
        --cool-knee: #2f8f83;
        --cool-shoulder: #4773a8;
    }

    .stApp {
        background:
            radial-gradient(circle at top left, rgba(120, 173, 219, 0.22), transparent 30%),
            linear-gradient(180deg, #f4f9fd 0%, #e8f1f9 100%);
        color: var(--cool-ink);
    }

    .block-container {
        padding-top: 2rem;
        padding-bottom: 2rem;
    }

    h1, h2, h3 {
        color: var(--cool-ink);
        letter-spacing: -0.02em;
    }

    .cool-hero {
        background: linear-gradient(135deg, rgba(30, 75, 112, 0.95), rgba(61, 126, 170, 0.88));
        border: 1px solid rgba(255, 255, 255, 0.2);
        border-radius: 18px;
        padding: 1.2rem 1.4rem;
        margin-bottom: 1rem;
        color: #f5fbff;
        box-shadow: 0 16px 36px rgba(40, 83, 122, 0.16);
    }

    .cool-caption {
        color: rgba(245, 251, 255, 0.82);
        font-size: 0.95rem;
    }

    div[data-testid="metric-container"] {
        background: var(--cool-surface);
        border: 1px solid var(--cool-border);
        border-radius: 16px;
        padding: 0.9rem 1rem;
        box-shadow: 0 10px 24px rgba(58, 91, 122, 0.08);
    }

    div[data-testid="metric-container"] label,
    div[data-testid="metric-container"] [data-testid="stMetricLabel"] {
        color: var(--cool-muted);
    }

    div[data-testid="stDataFrame"] {
        background: var(--cool-surface);
        border: 1px solid var(--cool-border);
        border-radius: 16px;
        padding: 0.3rem;
        box-shadow: 0 10px 24px rgba(58, 91, 122, 0.06);
    }

    .cool-panel {
        background: var(--cool-surface);
        border: 1px solid var(--cool-border);
        border-radius: 18px;
        padding: 1rem 1.1rem;
        box-shadow: 0 12px 28px rgba(58, 91, 122, 0.08);
        margin-bottom: 1rem;
    }

    .cool-knee {
        border-left: 6px solid var(--cool-knee);
    }

    .cool-shoulder {
        border-left: 6px solid var(--cool-shoulder);
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="cool-hero">
        <h1 style="margin:0;">Knee/Shoulder Stock Monitor</h1>
        <div class="cool-caption">Daily close-based reversal monitoring dashboard for Korean stocks.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

config = load_config()
paths = config["paths"]


def load_latest_signals(signal_dir: str) -> tuple[pd.DataFrame, str | None]:
    files = sorted(Path(signal_dir).glob("*_signals.csv"))
    if not files:
        return pd.DataFrame(), None
    latest = files[-1]
    return pd.read_csv(latest, dtype={"symbol": str}), latest.stem.replace("_signals", "")


signals_df, signal_date = load_latest_signals(paths["signal_dir"])
validation_df = load_validation_history(Path(paths["validation_file"]))

if signals_df.empty:
    st.warning("No signal file found yet. Run `python3 run_daily.py` first.")
    st.stop()

header_cols = st.columns(4)
header_cols[0].metric("Signal Date", signal_date or "-")
header_cols[1].metric("Knee Strong", int((signals_df["knee_grade"] == "Strong").sum()))
header_cols[2].metric("Shoulder Strong", int((signals_df["shoulder_grade"] == "Strong").sum()))
header_cols[3].metric("Symbols", len(signals_df))

left, right = st.columns(2)

with left:
    st.markdown('<div class="cool-panel cool-knee">', unsafe_allow_html=True)
    st.subheader("Knee Candidates")
    knee_view = signals_df[signals_df["knee_score"] >= config["runtime"]["signal_threshold"]].copy()
    st.dataframe(
        knee_view[["symbol", "name", "close", "pct_change", "knee_score", "knee_grade", "vol_ratio_20", "knee_reasons"]],
        use_container_width=True,
        hide_index=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

with right:
    st.markdown('<div class="cool-panel cool-shoulder">', unsafe_allow_html=True)
    st.subheader("Shoulder Candidates")
    shoulder_view = signals_df[signals_df["shoulder_score"] >= config["runtime"]["signal_threshold"]].copy()
    st.dataframe(
        shoulder_view[["symbol", "name", "close", "pct_change", "shoulder_score", "shoulder_grade", "vol_ratio_20", "shoulder_reasons"]],
        use_container_width=True,
        hide_index=True,
    )
    st.markdown("</div>", unsafe_allow_html=True)

st.markdown('<div class="cool-panel">', unsafe_allow_html=True)
st.subheader("Symbol Detail")
symbol = st.selectbox("Symbol", signals_df["symbol"] + " | " + signals_df["name"])
selected_symbol = symbol.split(" | ", 1)[0]
selected_row = signals_df[signals_df["symbol"] == selected_symbol].iloc[0]
history = load_existing_history(Path(paths["raw_dir"]) / f"{selected_symbol}.csv")

if not history.empty:
    figure = go.Figure()
    figure.add_trace(go.Scatter(x=history["date"], y=history["close"], mode="lines", name="Close"))
    if "ma_20" in history.columns:
        figure.add_trace(go.Scatter(x=history["date"], y=history["ma_20"], mode="lines", name="MA20"))
    figure.update_layout(
        height=420,
        margin=dict(l=20, r=20, t=20, b=20),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(232, 241, 249, 0.35)",
        font=dict(color="#12324a"),
    )
    st.plotly_chart(figure, use_container_width=True)

detail_cols = st.columns(2)
detail_cols[0].write(
    {
        "knee_score": int(selected_row["knee_score"]),
        "knee_grade": selected_row["knee_grade"],
        "knee_reasons": selected_row["knee_reasons"],
        "knee_confirmed": bool(selected_row["knee_confirmed"]),
    }
)
detail_cols[1].write(
    {
        "shoulder_score": int(selected_row["shoulder_score"]),
        "shoulder_grade": selected_row["shoulder_grade"],
        "shoulder_reasons": selected_row["shoulder_reasons"],
        "shoulder_confirmed": bool(selected_row["shoulder_confirmed"]),
    }
)
st.markdown("</div>", unsafe_allow_html=True)

st.markdown('<div class="cool-panel">', unsafe_allow_html=True)
st.subheader("Validation")
if not validation_df.empty:
    symbol_validation = validation_df[validation_df["symbol"] == selected_symbol]
    st.dataframe(symbol_validation, use_container_width=True, hide_index=True)
else:
    st.info("Validation data will appear after enough forward days have accumulated.")
st.markdown("</div>", unsafe_allow_html=True)
