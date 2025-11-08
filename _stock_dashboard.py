import streamlit as st
import subprocess
import sys
import time
import pandas as pd
import openpyxl
from pathlib import Path
import os

st.set_page_config(
    page_title="주식 데이터 대시보드",
    page_icon="📈",
    layout="wide"
)

st.title("📈 주식 데이터 대시보드")
st.markdown("---")

# 사이드바에 실행 버튼
with st.sidebar:
    st.header("데이터 업데이트")
    if st.button("🔄 데이터 갱신 시작", type="primary", use_container_width=True):
        st.session_state.run_update = True

# 메인 영역
if 'run_update' not in st.session_state:
    st.session_state.run_update = False
    
if 'data_loaded' not in st.session_state:
    st.session_state.data_loaded = False

if st.session_state.run_update:
    with st.sidebar:
        st.subheader("진행 상황")
        progress_bar = st.progress(0)
        status_text = st.empty()
    
    scripts = [
        ("_totalS.py", "S20/S60/S120 계산"),
        ("_totalZ.py", "Z20/Z60/Z120 계산"),
        ("_gap.py", "GAP 계산")
    ]
    results = []
    for idx, (script, description) in enumerate(scripts):
        with st.sidebar:
            status_text.text(f"⏳ {description} 중... ({idx+1}/{len(scripts)})")
        try:
            result = subprocess.run(
                [sys.executable, script],
                capture_output=True,
                text=True,
                timeout=300
            )
            if result.returncode == 0:
                results.append({
                    'script': script,
                    'description': description,
                    'status': '✅ 성공',
                    'output': result.stdout
                })
                with st.sidebar:
                    st.success(f"✅ {description} 완료!")
            else:
                results.append({
                    'script': script,
                    'description': description,
                    'status': '❌ 실패',
                    'output': result.stderr
                })
                with st.sidebar:
                    st.error(f"❌ {description} 실패!")
        except subprocess.TimeoutExpired:
            results.append({
                'script': script,
                'description': description,
                'status': '⏱️ 타임아웃',
                'output': '스크립트 실행 시간 초과 (5분)'
            })
            with st.sidebar:
                st.error(f"⏱️ {description} 타임아웃!")
        except Exception as e:
            results.append({
                'script': script,
                'description': description,
                'status': '❌ 오류',
                'output': str(e)
            })
            with st.sidebar:
                st.error(f"❌ {description} 오류: {str(e)}")
        with st.sidebar:
            progress_bar.progress((idx + 1) / len(scripts))
        time.sleep(0.5)
    with st.sidebar:
        status_text.text("✅ 모든 데이터 갱신 완료!")
        st.balloons()
        st.markdown("---")
        st.subheader("📊 실행 결과 요약")
        for result in results:
            with st.expander(f"{result['status']} {result['description']}", expanded=False):
                st.code(result['output'], language='text')
        if st.button("🔄 다시 실행"):
            st.session_state.run_update = True
            st.rerun()
    
    # 데이터 갱신 완료 후 메인 화면 표시 플래그 설정
    st.session_state.data_loaded = True
    st.session_state.run_update = False
    st.rerun()  # 메인 화면 표시를 위해 재실행

# 메인 화면 - 데이터 갱신 완료 후에만 표시
if st.session_state.data_loaded:
    # 메인 화면 - 최신 데이터 표시
    st.header("📊 종목별 최신 지표 데이터")
    
    # 엑셀 파일 찾기
    excel_files = list(Path('.').glob('_stock_value.xlsx'))
    
    if excel_files:
        excel_file = excel_files[0]
        
        try:
            # 엑셀 파일에서 데이터 읽기
            wb = openpyxl.load_workbook(excel_file, data_only=True)
            
            # '종목' 시트에서 종목코드와 종목명 매핑 가져오기
            stock_info = {}  # {종목코드: 종목명}
            if '종목' in wb.sheetnames:
                ws_stock = wb['종목']
                for row in ws_stock.iter_rows(min_row=2, max_col=2):  # 2행부터 2개 컬럼
                    stock_name = row[0].value  # A열: 종목명
                    stock_code = row[1].value  # B열: 종목코드
                    if stock_code and stock_name:
                        stock_info[stock_code] = stock_name
            
            # 각 시트별 최신 데이터 수집
            sheet_names = ['z20', 'z60', 'z120', 's20', 's60', 's120', 'gap']
            data_dict = {}
            latest_date = None  # 최신 날짜 저장
            
            # 종목코드로 데이터 딕셔너리 초기화
            for stock_code, stock_name in stock_info.items():
                data_dict[stock_code] = {
                    '종목코드': stock_code,
                    '종목명': stock_name
                }
            
            for sheet_name in sheet_names:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    
                    max_row = ws.max_row
                    max_col = ws.max_column
                    
                    # 최신 날짜 가져오기 (첫 번째 시트에서만, 헤더 행의 마지막 값이 있는 컬럼)
                    if latest_date is None and max_col > 2:
                        for col_idx in range(max_col, 2, -1):
                            date_value = ws.cell(row=1, column=col_idx).value
                            if date_value is not None and date_value != '':
                                latest_date = date_value
                                break
                    
                    # 각 행(종목)을 순회하며 최신 값 가져오기
                    for row_idx in range(2, max_row + 1):  # 2행부터 (1행은 헤더)
                        stock_code = ws.cell(row=row_idx, column=2).value  # 두 번째 컬럼이 종목코드
                        
                        if stock_code and stock_code in data_dict:
                            # 뒤에서부터 값이 있는 컬럼 찾기 (3번째 컬럼부터 시작, 1열=종목명, 2열=종목코드)
                            value = None
                            for col_idx in range(max_col, 2, -1):  # 마지막 컬럼부터 3번째 컬럼까지
                                cell_value = ws.cell(row=row_idx, column=col_idx).value
                                if cell_value is not None and cell_value != '':
                                    value = cell_value
                                    break
                            
                            data_dict[stock_code][sheet_name.upper()] = value if value is not None else '-'
            
            wb.close()
            
            if data_dict:
                # DataFrame 생성
                df = pd.DataFrame.from_dict(data_dict, orient='index')
                df = df.reset_index(drop=True)
                
                # 컬럼 순서 정리 (종목코드, 종목명, 나머지 지표)
                column_order = ['종목코드', '종목명', 'Z20', 'Z60', 'Z120', 'S20', 'S60', 'S120', 'GAP']
                existing_columns = [col for col in column_order if col in df.columns]
                df = df[existing_columns]
                
                # 필터링 옵션
                st.markdown("### 🔍 필터 옵션")
                col1, col2 = st.columns(2)
                
                with col1:
                    search_stock = st.text_input("🔎 종목명/종목코드 검색", placeholder="종목명 또는 종목코드를 입력하세요")
                
                with col2:
                    sort_by = st.selectbox(
                        "정렬 기준",
                        options=['종목코드', '종목명'] + [col for col in df.columns if col not in ['종목코드', '종목명']],
                        index=0
                    )
                
                # 검색 필터 적용 (종목명 또는 종목코드로 검색)
                if search_stock:
                    df_filtered = df[
                        df['종목명'].str.contains(search_stock, case=False, na=False) |
                        df['종목코드'].astype(str).str.contains(search_stock, case=False, na=False)
                    ]
                else:
                    df_filtered = df.copy()
                
                # 정렬
                if sort_by not in ['종목코드', '종목명']:
                    # 숫자형으로 변환 후 정렬
                    df_filtered[sort_by] = pd.to_numeric(df_filtered[sort_by], errors='coerce')
                    df_filtered = df_filtered.sort_values(by=sort_by, ascending=False)
                else:
                    df_filtered = df_filtered.sort_values(by=sort_by)
                
                # 데이터 표시
                st.markdown(f"### 📈 최신 데이터 ({len(df_filtered)}개 종목)")
                
                # 최신 날짜 표시
                if latest_date:
                    st.info(f"📅 데이터 기준일: **{latest_date}**")
                
                # 스타일링된 데이터프레임 표시
                st.dataframe(
                    df_filtered,
                    use_container_width=True,
                    height=600,
                    hide_index=True,
                    column_config={
                        "종목코드": st.column_config.TextColumn("종목코드", width="small"),
                        "종목명": st.column_config.TextColumn("종목명", width="small"),
                        "Z20": st.column_config.NumberColumn("Z20", format="%.2f", width="small"),
                        "Z60": st.column_config.NumberColumn("Z60", format="%.2f", width="small"),
                        "Z120": st.column_config.NumberColumn("Z120", format="%.2f", width="small"),
                        "S20": st.column_config.NumberColumn("S20", format="%.2f", width="small"),
                        "S60": st.column_config.NumberColumn("S60", format="%.2f", width="small"),
                        "S120": st.column_config.NumberColumn("S120", format="%.2f", width="small"),
                        "GAP": st.column_config.NumberColumn("GAP", format="%.2f", width="small"),
                    }
                )
                
                # ...existing code...
                
            else:
                st.warning("⚠️ 데이터를 찾을 수 없습니다. 먼저 데이터를 갱신해 주세요.")
                
        except Exception as e:
            st.error(f"❌ 데이터 로딩 오류: {str(e)}")
    else:
        st.warning("⚠️ _stock_value.xlsx 파일을 찾을 수 없습니다.")

else:
    # 초기 화면 - 데이터 갱신 전
    st.info("👈 왼쪽 사이드바에서 '데이터 갱신 시작' 버튼을 클릭하여 데이터를 먼저 로드하세요.")

st.markdown("---")
st.caption("📂 파일: _stock_value.xlsx")
