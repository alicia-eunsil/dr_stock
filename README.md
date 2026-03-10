# Knee/Shoulder Stock Monitor

한국투자증권 API를 사용해 매일 오후 4시에 종목별 일봉을 수집하고, `무릎(knee)` / `어깨(shoulder)` 후보를 점수화해 보여주는 프로젝트입니다.

## 구성

- `run_daily.py`: 배치 실행 진입점
- `app.py`: Streamlit 대시보드
- `src/knee_shoulder/`: API, 마스터, 지표, 신호, 검증 로직
- `data/master/stocks_kr.csv`: 종목 마스터
- `data/raw/`: 종목별 누적 일봉 CSV
- `data/patches/`: 일자별 패치 CSV
- `data/signals/`: 일자별 신호 CSV
- `data/validation/signal_validation.csv`: 누적 검증 결과

## 시작

1. `config.example.json`을 `config.json`으로 복사
2. 민감정보는 파일 대신 환경변수로 설정
3. 종목 마스터 생성

```bash
export KIS_APP_KEY='YOUR_APP_KEY'
export KIS_APP_SECRET='YOUR_APP_SECRET'
export KIS_BASE_URL='https://openapi.koreainvestment.com:9443'

python3 run_daily.py \
  --rebuild-master \
  --master-source '/Users/alicia/Desktop/#python/#Version2_broad_/KR_Stocks_Individual.xlsx'
```

4. 일일 배치 실행

```bash
python3 run_daily.py
```

5. 대시보드 실행

```bash
streamlit run app.py
```

## launchd 자동 실행

`launchd`는 터미널의 `export` 값을 자동으로 가져오지 않으므로, 아래 값을 `~/.bash_profile`에 넣어둬야 합니다.

```bash
export KIS_APP_KEY='YOUR_APP_KEY'
export KIS_APP_SECRET='YOUR_APP_SECRET'
export KIS_BASE_URL='https://openapi.koreainvestment.com:9443'
```

자동 실행 스크립트:

- [run_daily_launchd.sh](/Users/alicia/Desktop/#python/knee_shoulder_stock/deploy/launchd/run_daily_launchd.sh)

`launchd` 설정 파일:

- [com.alicia.knee-shoulder-stock.plist](/Users/alicia/Desktop/#python/knee_shoulder_stock/deploy/launchd/com.alicia.knee-shoulder-stock.plist)

## 비고

- 종목 목록 원본은 `KR_Stocks_Individual.xlsx`의 `종목` 시트를 사용합니다.
- `data/raw/`와 `data/signals/` 등 실행 산출물은 `.gitignore`에 포함했습니다.
- 민감정보는 기본적으로 `KIS_APP_KEY`, `KIS_APP_SECRET`, `KIS_BASE_URL` 환경변수에서 읽습니다.
- `secrets.json`은 로컬 fallback 용도이며 GitHub에 올리면 안 됩니다.
- 첫 실행은 `history_lookback_days`만큼 넓게 적재하고, 이후 실행은 각 종목의 최신 저장일 기준 `incremental_recheck_days`만큼만 재조회합니다.
- 첫 버전은 가격/거래량 기반 신호에 집중했고, 투자자별 매매량 API는 2차 확장용으로 남겨두었습니다.
