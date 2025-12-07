# _stock_dashboard.py
# 추가: 원자료 보기, 더보기 기능 추가
# 추가: 지표별 탭 추가, 탭별 기능 함수화
# 추가: 종합탭 내 지수 보여주기 
# 추가: 종가를 활용하여 표준편차 구하고 값을 보여주는 기능 추가

import streamlit as st
import subprocess
import sys
import pandas as pd
import openpyxl
from pathlib import Path
import bcrypt
from datetime import datetime, date, timedelta

# ======================================
# 페이지 설정
# ======================================
st.set_page_config(page_title="주식 데이터 대시보드", page_icon="📈", layout="wide")

# ======================================
# 0. 인증 (간단 비밀번호)
# ======================================
ACCESS_CODE_HASH = b"$2b$12$gDBpQYK.g938H.8cNwLeUu/VRidCP1GxqusJiEQzVnvaSrG4CBE6K"

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False

if not st.session_state["authenticated"]:
    st.title("🔒 Access Required")
    st.write("Please enter the access code to open the dashboard.")

    with st.form("auth_form"):
        code = st.text_input("Enter access code", type="password")
        submitted = st.form_submit_button("Submit")

    if submitted:
        if bcrypt.checkpw(code.encode(), ACCESS_CODE_HASH):
            st.session_state["authenticated"] = True
            st.success("Access granted")
            st.rerun()
        else:
            st.error("Invalid code")

    st.stop()

# ======================================
# 1. 전역 상태 변수
# ======================================
if "run_update" not in st.session_state:
    st.session_state.run_update = False
if "data_loaded" not in st.session_state:
    st.session_state.data_loaded = True

# 🔥 종합 탭 날짜 확장용
if "show_days" not in st.session_state:
    st.session_state.show_days = 10  # 시작: 최근 10일

# 🔥 원자료 탭 날짜 확장용
if "show_days_raw" not in st.session_state:
    st.session_state.show_days_raw = 10  # 시작: 최근 10일

# ======================================
# 2. 날짜/포맷 유틸 함수
# ======================================
def _to_datetime(v):
    """엑셀/문자열/숫자 등 다양한 형태의 날짜를 datetime으로 변환"""
    if isinstance(v, (datetime, date)):
        return datetime(v.year, v.month, v.day)

    if isinstance(v, (int, float)):
        base = datetime(1899, 12, 30)
        try:
            return base + timedelta(days=int(v))
        except:
            return None

    s = str(v).strip()
    if not s:
        return None

    for fmt in ("%Y-%m-%d", "%Y.%m.%d.", "%Y.%m.%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(s, fmt)
        except:
            pass

    digits = "".join(ch for ch in s if ch.isdigit())
    if len(digits) == 8:
        try:
            return datetime.strptime(digits, "%Y%m%d")
        except:
            pass

    return None


def format_excel_date(v):
    """_to_datetime로 바꾼 날짜를 YYYY.MM.DD. 형식 문자열로 변환"""
    dt = _to_datetime(v)
    if dt:
        return dt.strftime("%Y.%m.%d.")
    s = str(v)
    s = s.replace("-", ".").replace("/", ".")
    if not s.endswith("."):
        s += "."
    return s


def _format_z_cell(v):
    val = pd.to_numeric(v, errors="coerce")
    if pd.isna(val):
        return "-"
    out = f"{val:.0f}"
    if val > 100:
        out += " 🔴"
    elif val < -100:
        out += " 🔵"
    return out


def _format_s_cell(v):
    val = pd.to_numeric(v, errors="coerce")
    if pd.isna(val):
        return "-"
    out = f"{val:.0f}"
    if abs(val - 100) < 0.1:
        out += " 🔴"
    elif abs(val - 0) < 0.1:
        out += " 🔵"
    return out


def _format_q_cell(v):
    val = pd.to_numeric(v, errors="coerce")
    if pd.isna(val):
        return "-"
    out = f"{val:.0f}"
    if val > 100:
        out += " 🔴"
    elif val < 25:
        out += " 🔵"
    return out

# ======================================
# 3. 뷰 렌더링 함수들
# ======================================
def render_total_view(indicator_df, selected_labels, indicator_range_msg, total_days, index_df=None):
    """
    1️⃣ 종합 탭
    - 멀티헤더(날짜×지표) 구조
    - 맨 아래 평균 행
    - 그 아래 KOSPI/KOSDAQ/KOSPI200 행 추가
    """
    if indicator_df is None:
        st.warning("⚠️ 종합 데이터를 불러올 수 없습니다.")
        return

    st.markdown("### 🔍 필터 옵션 (종합)")
    c1, c2 = st.columns(2)
    with c1:
        search = st.text_input("🔎 종목명/종목코드 검색", key="search_total")
    with c2:
        sort_by = st.selectbox("정렬 기준", ["종목코드", "종목명"], key="sort_total")

    # 검색 적용
    df_f = indicator_df.copy()
    if search:
        df_f = df_f[
            df_f["종목명"].astype(str).str.contains(search, case=False) |
            df_f["종목코드"].astype(str).str.contains(search, case=False)
        ]

    df_f = df_f.sort_values(by=sort_by)

    st.info(indicator_range_msg)

    # --------------------------------------
    # 🔥 멀티헤더 생성 (1행: 날짜, 2행: 지표명)
    # --------------------------------------
    metrics = ["Z20", "Z60", "Z120", "S20", "S60", "S120", "GAP", "QUANT", "STD"]
    base_cols = ["종목코드", "종목명"]
    df_show = df_f[base_cols].copy()

    col_tuples = [("", "종목코드"), ("", "종목명")]

    # 날짜 × 지표 조합 생성 (값 없으면 '-' 처리)
    for lbl in selected_labels:
        for m in metrics:
            key = (lbl, m)
            if key in df_f.columns:
                df_show[(lbl, m)] = df_f[key]
            else:
                df_show[(lbl, m)] = "-"
            col_tuples.append((lbl, m))

    df_show.columns = pd.MultiIndex.from_tuples(col_tuples)

    # --------------------------------------
    # 🔥 평균 행 추가 (맨 마지막 행)
    # --------------------------------------
    avg_row = []
    for col in df_show.columns:
        if col == ("", "종목코드"):
            avg_row.append("AVG")
        elif col == ("", "종목명"):
            avg_row.append("평균")
        else:
            lbl, m = col
            key = (lbl, m)
            if key in df_f.columns:
                s = pd.to_numeric(df_f[key], errors="coerce")
                avg_val = s.mean(skipna=True)
                avg_row.append(f"{avg_val:.2f}")
            else:
                avg_row.append(None)

    df_show.loc[len(df_show)] = avg_row  # 평균 행 추가

    # --------------------------------------
    # Z/S/Q/GAP 포맷 이모지 적용
    # --------------------------------------
    for lbl in selected_labels:
        for m in ["Z20", "Z60", "Z120"]:
            col = (lbl, m)
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(_format_z_cell)

        for m in ["S20", "S60", "S120"]:
            col = (lbl, m)
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(_format_s_cell)

        for metric in ["GAP", "STD"]:
            col = (lbl, metric)
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(
                    lambda v: "-" if pd.isna(pd.to_numeric(v, errors="coerce")) else v
                )

        for m in ["QUANT"]:
            col = (lbl, m)
            if col in df_show.columns:
                df_show[col] = df_show[col].apply(_format_q_cell)

    # --------------------------------------
    # 🔽 지수(KOSPI/KOSDAQ/KOSPI200) 행 추가
    #   - 같은 날짜(lbl)에 대해 "첫 번째 컬럼"에만 값 넣고,
    #     나머지 7개 지표 칸은 공백으로 둠
    # --------------------------------------
    if index_df is not None and not index_df.empty:
        for _, idx_row in index_df.iterrows():
            new_row_vals = []
            used_dates = set()  # 같은 날짜에 한 번만 값 넣기 위한 기록

            for col in df_show.columns:
                if col == ("", "종목코드"):
                    new_row_vals.append(idx_row.get("업종코드", ""))
                elif col == ("", "종목명"):
                    new_row_vals.append(idx_row.get("업종명", ""))
                else:
                    lbl, m = col  # lbl: 날짜 라벨, m: Z20/S20/...

                    if lbl not in used_dates:
                        val = idx_row.get(lbl, None)
                        new_row_vals.append(val if pd.notna(val) else "")
                        used_dates.add(lbl)
                    else:
                        new_row_vals.append("")

            df_show.loc[len(df_show)] = new_row_vals  # 지수 행 추가

    # 인덱스 설정 (종목코드·종목명)
    df_show = df_show.set_index([("", "종목코드"), ("", "종목명")])

    st.dataframe(
        df_show,
        use_container_width=True,
        height=600,
    )

    # 🔥 과거 확장 버튼 (종합)
    if st.button("⬅ 과거 10일 더보기(종합)", disabled=(total_days <= st.session_state.show_days)):
        st.session_state.show_days = min(st.session_state.show_days + 10, total_days)
        st.rerun()


def render_metric_view(indicator_df, selected_labels):
    """
    2️⃣ 지표별 탭:
    - 1열: 종목코드
    - 2열: 종목명
    - 이후: 날짜별 선택 지표값
    """
    st.subheader("📈 지표 선택")

    if indicator_df is None or len(indicator_df) == 0:
        st.warning("⚠️ 지표별 데이터를 불러올 수 없습니다.")
        return

    metric_options = ["Z20", "Z60", "Z120",
                      "S20", "S60", "S120",
                      "GAP", "QUANT", "STD"]

    # 실제 존재하는 지표만
    available = []
    for m in metric_options:
        if any(((lbl, m) in indicator_df.columns) for lbl in selected_labels):
            available.append(m)

    if not available:
        st.error("indicator_df에 S/Z/GAP/QUANT/STD 관련 컬럼이 없습니다.")
        st.write("현재 indicator_df.columns 예시:", list(indicator_df.columns)[:20])
        return

    metric = st.selectbox("지표를 선택하세요", available, index=0)

    # -------------------------
    # DF 구성 (종목코드, 종목명 + 날짜별 값)
    # -------------------------
    df_metric = indicator_df[["종목코드", "종목명"]].copy()

    for lbl in selected_labels:
        col_key = (lbl, metric)
        if col_key in indicator_df.columns:
            df_metric[lbl] = indicator_df[col_key]
        else:
            df_metric[lbl] = None

    # 값 포맷팅
    def _format_plain(v):
        val = pd.to_numeric(v, errors="coerce")
        if pd.isna(val):
            return "-"
        return f"{val:.0f}"

    def _format_std_cell(v):
        val = pd.to_numeric(v, errors="coerce")
        if pd.isna(val):
            return "-"
        return f"{val:.2f}"

    if metric == "STD":
        formatter = _format_std_cell
    elif metric.startswith("S"):
        formatter = _format_s_cell
    elif metric.startswith("Z"):
        formatter = _format_z_cell
    else:
        formatter = _format_plain

    for lbl in selected_labels:
        if lbl in df_metric.columns:
            df_metric[lbl] = df_metric[lbl].apply(formatter)

    # 🔍 필터 + 정렬
    st.markdown("### 🔍 필터 옵션 (지표별)")
    c1, c2 = st.columns(2)
    with c1:
        search_metric = st.text_input("🔎 종목명/종목코드 검색", key="search_metric")
    with c2:
        sort_metric = st.selectbox("정렬 기준", ["종목코드", "종목명"], key="sort_metric")

    df_filtered = df_metric.copy()
    if search_metric:
        df_filtered = df_filtered[
            df_filtered["종목명"].astype(str).str.contains(search_metric, case=False)
            | df_filtered["종목코드"].astype(str).str.contains(search_metric, case=False)
        ]

    df_filtered = df_filtered.sort_values(by=sort_metric).reset_index(drop=True)

    # 날짜 범위 안내
    if selected_labels:
        oldest_label = selected_labels[0]
        latest_label = selected_labels[-1]
        st.info(
            f"📅 지표별 표시 범위: **{oldest_label} ~ {latest_label}** "
            f"(최근 {len(selected_labels)}일)"
        )

    # 테이블 출력
    st.markdown(f"### 📋 {metric} · 추이")

    column_config = {
        "종목코드": st.column_config.TextColumn("종목코드", width="small", pinned="left"),
        "종목명": st.column_config.TextColumn("종목명", width="small", pinned="left"),
    }
    for lbl in selected_labels:
        if lbl in df_filtered.columns:
            column_config[lbl] = st.column_config.TextColumn(lbl)

    st.dataframe(
        df_filtered,
        use_container_width=True,
        height=600,
        hide_index=True,
        column_config=column_config,
    )

    # 🔥 과거 확장 버튼 (지표별)
    global total_days
    if st.button("⬅ 과거 10일 더보기(지표별)", disabled=(total_days <= st.session_state.show_days)):
        st.session_state.show_days = min(st.session_state.show_days + 10, total_days)
        st.rerun()


def render_raw_view(close_df, close_range_msg, total_close_days):
    """
    3️⃣ 원자료(종가) 탭
    - 종목코드/종목명 + 날짜별 종가
    """
    if close_df is None:
        st.warning("⚠️ 원자료(종가) 데이터를 불러올 수 없습니다.")
        return

    st.markdown("### 🔍 필터 옵션 (원자료)")
    r1, r2 = st.columns(2)
    with r1:
        search_raw = st.text_input("🔎 종목명/종목코드 검색", key="search_raw")
    with r2:
        sort_raw = st.selectbox("정렬 기준", ["종목코드", "종목명"], key="sort_raw")

    df_raw = close_df.copy()

    if search_raw:
        df_raw = df_raw[
            df_raw["종목코드"].astype(str).str.contains(search_raw, case=False) |
            df_raw["종목명"].astype(str).str.contains(search_raw, case=False)
        ]

    df_raw = df_raw.sort_values(by=sort_raw)

    st.info(close_range_msg)

    # 날짜 컬럼 추출
    date_cols = [c for c in df_raw.columns if c not in ["종목코드", "종목명"]]

    # 컬럼 순서 고정
    df_raw = df_raw[["종목코드", "종목명"] + date_cols]

    # 날짜 컬럼 숫자 변환
    for c in date_cols:
        df_raw[c] = pd.to_numeric(df_raw[c], errors="coerce")

    column_config = {
        "종목코드": st.column_config.TextColumn("종목코드", width="small", pinned="left"),
        "종목명": st.column_config.TextColumn("종목명", width="small", pinned="left"),
    }
    for c in date_cols:
        column_config[c] = st.column_config.NumberColumn(
            c,
            format="%.0f",
        )

    st.dataframe(
        df_raw,
        use_container_width=True,
        height=600,
        hide_index=True,
        column_config=column_config,
    )

    # 🔥 과거 확장 버튼 (원자료)
    if st.button("⬅ 과거 10일 더보기(종가)", disabled=(total_close_days <= st.session_state.show_days_raw)):
        st.session_state.show_days_raw = min(st.session_state.show_days_raw + 10, total_close_days)
        st.rerun()

# ======================================
# 4. 사이드바: 데이터 갱신 버튼
# ======================================
with st.sidebar:
    excel_path = Path("_stock_value.xlsx")
    if excel_path.exists():
        with open(excel_path, "rb") as f:
            st.download_button(
                label="📥 최신 데이터 다운로드",
                data=f,
                file_name="_stock_value.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="download_excel",
            )

    if st.button("🔄 데이터 갱신 시작"):
        st.session_state.run_update = True

# ======================================
# 5. 데이터 갱신 실행 (외부 스크립트 호출)
# ======================================
if st.session_state.run_update:
    with st.sidebar:
        st.subheader("진행 상황")
        pb = st.progress(0)
        msg = st.empty()

    scripts = [
        ("_totalS.py", "S20/S60/S120 계산"),
        ("_totalZ.py", "Z20/Z60/Z120 계산"),
        ("_gap.py", "GAP 계산"),
        ("_quant.py", "거래량 계산"),
        ("_std.py", "표준편차 계산"),
    ]

    for idx, (sc, desc) in enumerate(scripts):
        msg.write(f"{desc} 실행 중...")
        try:
            result = subprocess.run(
                [sys.executable, sc], capture_output=True, text=True, timeout=300
            )
            if result.returncode == 0:
                st.sidebar.success(f"{desc} 완료")
            else:
                st.sidebar.error(f"{desc} 실패")
        except:
            st.sidebar.error(f"{desc} 오류 발생")

        pb.progress((idx + 1) / len(scripts))

    st.session_state.data_loaded = True
    st.session_state.run_update = False
    st.rerun()

# ======================================
# 6. 엑셀 파일 로드
# ======================================
excel_files = list(Path(".").glob("_stock_value.xlsx"))
if not excel_files:
    st.error("_stock_value.xlsx 파일을 찾지 못했습니다. "
             "GitHub Actions가 아직 안 돌았거나, "
             "필요하면 왼쪽의 '데이터 갱신 시작' 버튼을 눌러주세요.")
    st.stop()

excel_file = excel_files[0]
wb = openpyxl.load_workbook(excel_file, data_only=True)

# ======================================
# 7. 종목 정보 로딩 (종목 시트)
# ======================================
stock_info = {}
if "종목" in wb.sheetnames:
    ws = wb["종목"]
    for r in ws.iter_rows(min_row=2, max_col=2):
        name = r[0].value
        code = r[1].value
        if code and name:
            stock_info[code] = name

# ======================================
# 8. 종합(Z20/Z60/S/GAP/QUANT) 데이터 로딩
# ======================================
sheet_names = ["z20", "z60", "z120", "s20", "s60", "s120", "gap", "quant", "std"]

base_ws = None
for s in sheet_names:
    if s in wb.sheetnames:
        base_ws = wb[s]
        break

indicator_df = None
indicator_date_infos = []
total_days = 0
selected_labels = []
indicator_range_msg = ""

if base_ws:
    max_col = base_ws.max_column

    # 기준 시트에서 날짜 헤더 수집 (1행, 3열~)
    for col in range(3, max_col + 1):
        raw = base_ws.cell(row=1, column=col).value
        if raw is None:
            continue
        dt = _to_datetime(raw)
        label = format_excel_date(raw)
        indicator_date_infos.append((col, raw, dt, label))

    indicator_date_infos = sorted(
        indicator_date_infos,
        key=lambda x: (x[2] is None, x[2] or datetime.min)
    )

    total_days = len(indicator_date_infos)

    show_days = min(st.session_state.show_days, total_days)
    start_idx = total_days - show_days
    selected_infos = indicator_date_infos[start_idx:]
    selected_labels = [lbl for _, _, _, lbl in selected_infos]

    oldest_label = selected_infos[0][3]
    latest_label = selected_infos[-1][3]
    indicator_range_msg = (
        f"📅 종합 표시 범위: **{oldest_label} ~ {latest_label}** "
        f"(최근 {show_days}일 / 전체 {total_days}일)"
    )

    # 종목별 데이터 딕셔너리 구성
    data_dict = {code: {"종목코드": code, "종목명": name}
                 for code, name in stock_info.items()}

    # 시트별 데이터 채우기 (날짜 문자열 기준 매칭)
    for s in sheet_names:
        if s not in wb.sheetnames:
            continue

        ws = wb[s]
        max_row_s = ws.max_row
        max_col_s = ws.max_column

        label_to_col = {}
        for col in range(3, max_col_s + 1):
            raw = ws.cell(row=1, column=col).value
            if raw is None:
                continue
            lbl = format_excel_date(raw)
            label_to_col[lbl] = col

        for r in range(2, max_row_s + 1):
            code = ws.cell(row=r, column=2).value
            if code not in data_dict:
                continue

            for lbl in selected_labels:
                col_idx = label_to_col.get(lbl)
                if col_idx is None:
                    val = None
                else:
                    val = ws.cell(row=r, column=col_idx).value

                data_dict[code][(lbl, s.upper())] = val

    indicator_df = pd.DataFrame.from_dict(data_dict, orient="index").reset_index(drop=True)
else:
    indicator_df = None

# ======================================
# 9. 원자료(종가) 데이터 로딩
# ======================================
close_df = None
close_date_infos = []
total_close_days = 0
close_range_msg = ""

if "종가" in wb.sheetnames:
    ws = wb["종가"]
    max_col_c = ws.max_column

    # 날짜 헤더
    for col in range(3, max_col_c + 1):
        raw = ws.cell(row=1, column=col).value
        if raw is None:
            continue

        dt = _to_datetime(raw)

        if dt is None:
            digits = "".join(ch for ch in str(raw) if ch.isdigit())
            if len(digits) == 8:
                dt = datetime.strptime(digits, "%Y%m%d")

        if dt is None:
            continue

        label = dt.strftime("%Y.%m.%d.")
        close_date_infos.append((col, raw, dt, label))

    close_date_infos = sorted(
        close_date_infos,
        key=lambda x: (x[2] is None, x[2] or datetime.min)
    )

    total_close_days = len(close_date_infos)

    show_raw = min(st.session_state.show_days_raw, total_close_days)
    start_idx = total_close_days - show_raw
    selected_close_infos = close_date_infos[start_idx:]

    oldest_label = selected_close_infos[0][3]
    latest_label = selected_close_infos[-1][3]

    close_range_msg = (
        f"📅 종가 표시 범위: **{oldest_label} ~ {latest_label}** "
        f"(최근 {show_raw}일 / 전체 {total_close_days}일)"
    )

    close_dict = {code: {"종목명": name, "종목코드": code}
                  for code, name in stock_info.items()}

    max_row_c = ws.max_row

    for r in range(2, max_row_c + 1):
        code = ws.cell(row=r, column=2).value
        if code not in close_dict:
            continue

        for col_idx, raw, dt, label in selected_close_infos:
            val = ws.cell(row=r, column=col_idx).value
            close_dict[code][label] = val

    close_df = pd.DataFrame.from_dict(close_dict, orient="index").reset_index(drop=True)

    # 컬럼 이름을 yyyy.mm.dd. 형식으로 통일
    rename_map = {}
    for col in close_df.columns:
        if col in ["종목코드", "종목명"]:
            continue
        rename_map[col] = format_excel_date(col)

    close_df = close_df.rename(columns=rename_map)

# ======================================
# 10. 지수(KOSPI/KOSDAQ/KOSPI200) 데이터 로딩
#     → 종합 탭에서 selected_labels에 맞춰 사용할 index_df
# ======================================
index_df = None

if "지수" in wb.sheetnames and indicator_df is not None and selected_labels:
    ws_idx = wb["지수"]
    max_col_i = ws_idx.max_column

    index_date_infos = []
    for col in range(3, max_col_i + 1):
        raw = ws_idx.cell(row=1, column=col).value
        if raw is None:
            continue

        dt = _to_datetime(raw)
        if dt is None:
            digits = "".join(ch for ch in str(raw) if ch.isdigit())
            if len(digits) == 8:
                dt = datetime.strptime(digits, "%Y%m%d")
        if dt is None:
            continue

        label = dt.strftime("%Y.%m.%d.")
        index_date_infos.append((col, raw, dt, label))

    label_to_col_idx = {label: col for col, raw, dt, label in index_date_infos}

    index_rows = []
    max_row_i = ws_idx.max_row

    for r in range(2, max_row_i + 1):
        name = ws_idx.cell(row=r, column=1).value
        code = ws_idx.cell(row=r, column=2).value
        if not name or not code:
            continue

        row_dict = {
            "업종명": str(name),
            "업종코드": str(code),
        }

        for lbl in selected_labels:
            col_idx = label_to_col_idx.get(lbl)
            if col_idx is None:
                val = None
            else:
                val = ws_idx.cell(row=r, column=col_idx).value
            row_dict[lbl] = val

        index_rows.append(row_dict)

    if index_rows:
        index_df = pd.DataFrame(index_rows)

# ======================================
# 11. 엑셀 파일 닫기
# ======================================
wb.close()

# ======================================
# 12. 탭 구성 및 렌더링
# ======================================
tab_total, tab_metric, tab_raw = st.tabs(["1️⃣ 종합", "2️⃣ 지표별", "3️⃣ 원자료"])

with tab_total:
    if indicator_df is None:
        st.warning("⚠️ 종합 데이터를 불러올 수 없습니다.")
    else:
        render_total_view(
            indicator_df,
            selected_labels,
            indicator_range_msg,
            total_days,
            index_df=index_df,
        )

with tab_metric:
    if indicator_df is None:
        st.warning("⚠️ 지표별 데이터를 불러올 수 없습니다.")
    else:
        render_metric_view(indicator_df, selected_labels)

with tab_raw:
    if close_df is None:
        st.warning("⚠️ 원자료(종가) 데이터를 불러올 수 없습니다.")
    else:
        render_raw_view(close_df, close_range_msg, total_close_days)

st.markdown("---")
st.caption("Created by Alicia")
