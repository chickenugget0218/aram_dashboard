# 판매 대시보드 — 인수인계 문서

복지단 결산 데이터를 받아 자동으로 파싱·시각화하는 사내용 웹 대시보드입니다. Streamlit + Supabase로 만들어졌고, Streamlit Cloud에 배포돼 있습니다.

이 문서는 코드 구조와 각 함수의 역할을 인수인계가 가능하도록 정리한 것입니다.

---

## 목차

1. [전체 아키텍처](#1-전체-아키텍처)
2. [파일 구조](#2-파일-구조)
3. [데이터베이스 구조 (Supabase)](#3-데이터베이스-구조-supabase)
4. [app.py 코드 구조](#4-apppy-코드-구조)
   - [4-1. 임포트와 페이지 설정](#4-1-임포트와-페이지-설정)
   - [4-2. 인증 (check_password)](#4-2-인증-check_password)
   - [4-3. Supabase 연결 (get_supabase)](#4-3-supabase-연결-get_supabase)
   - [4-4. 고정 설정 (상수)](#4-4-고정-설정-상수)
   - [4-5. 유틸 함수 (logo_for, safe_path, company_code)](#4-5-유틸-함수)
   - [4-6. 파서 (parse_monthly, parse_daily)](#4-6-파서)
   - [4-7. DB 조회 헬퍼 (fetch_*)](#4-7-db-조회-헬퍼)
   - [4-8. 헤더 & 데이터 로드](#4-8-헤더--데이터-로드)
   - [4-9. 탭 1: 📊 대시보드](#4-9-탭-1--대시보드)
   - [4-10. 탭 2: 📁 파일 업로드](#4-10-탭-2--파일-업로드)
   - [4-11. 탭 3: 💰 매출 이력](#4-11-탭-3--매출-이력)
   - [4-12. 탭 4: 🎯 목표 설정](#4-12-탭-4--목표-설정)
5. [데이터 처리 규칙](#5-데이터-처리-규칙)
6. [운영 가이드](#6-운영-가이드)
7. [자주 막히는 지점](#7-자주-막히는-지점)

---

## 1. 전체 아키텍처

```
[사용자 브라우저]
       │
       ▼
[Streamlit Cloud]   ← 호스팅. GitHub repo 변경 감지하면 자동 재배포
       │
       │   파일 업로드 ──▶ [Supabase Storage]  ← 원본 xls/csv 보관
       │                          │
       │                     앱이 읽어 파싱
       │                          │
       └── 데이터 입출력 ───▶ [Supabase Postgres]
                                  ├── monthly   월별 판매금액
                                  ├── daily     일별 누적 매출
                                  ├── targets   목표 (LG생활건강 · 월별)
                                  └── memos     미납현황
```

핵심 정리:
- **사내 PC 불필요.** 브라우저만 있으면 누구나 URL로 접속.
- **앱 비밀번호로 보호.** Streamlit Cloud 무료 플랜은 퍼블릭 repo만 배포 가능하므로 코드는 공개되나, 키와 데이터는 Supabase에 격리.
- **데이터 흐름.** 사용자가 엑셀/CSV 업로드 → 앱이 파싱 → Supabase Postgres에 행 단위로 upsert → 대시보드가 다시 읽어 시각화.

---

## 2. 파일 구조

```
sales-dashboard/
├── app.py                   # 메인 앱 (한 파일)
├── requirements.txt         # 파이썬 패키지 목록
├── README.md                # 이 문서
├── .gitignore               # GitHub에 안 올릴 파일 목록
│
├── .streamlit/
│   └── secrets.toml         # Supabase 키 + 앱 비밀번호 (절대 GitHub에 안 올림)
│
└── logo/                    # 업체 로고 (한글 파일명)
    ├── LG생활건강.png
    ├── 비알코리아.png
    ├── 에너자이저.png
    ├── 라벨리.png
    ├── 메디카.png
    ├── 남양유업.png
    ├── 나사라.png
    ├── 삼립.png
    └── 티젠.png
```

**`requirements.txt`** 내용 (정확히 이 6줄만):
```
streamlit
pandas
altair
xlrd
openpyxl
supabase
```

**`.gitignore`** 내용:
```
.streamlit/secrets.toml
__pycache__/
*.pyc
.DS_Store
test_supabase.py
test_upload.py
```

**`.streamlit/secrets.toml`** 내용 (로컬 + Streamlit Cloud Secrets 양쪽에 동일):
```toml
SUPABASE_URL = "https://여러분프로젝트.supabase.co"
SUPABASE_KEY = "eyJhbGc..."
APP_PASSWORD = "사내공유비밀번호"
```

---

## 3. 데이터베이스 구조 (Supabase)

### Postgres 테이블 4개

```sql
CREATE TABLE monthly (
    id BIGSERIAL PRIMARY KEY,
    업체 TEXT NOT NULL,
    월 TEXT NOT NULL,                -- "2026-04"
    판매금액 BIGINT NOT NULL,
    업로드시각 TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(업체, 월)
);

CREATE TABLE daily (
    id BIGSERIAL PRIMARY KEY,
    업체 TEXT NOT NULL,
    날짜 TEXT NOT NULL,               -- "2026-06-17"
    매출액 BIGINT NOT NULL,
    업로드시각 TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(업체, 날짜)
);

CREATE TABLE targets (
    업체 TEXT NOT NULL,
    월 TEXT NOT NULL,
    목표금액 BIGINT NOT NULL,
    PRIMARY KEY(업체, 월)
);

CREATE TABLE memos (
    id BIGSERIAL PRIMARY KEY,
    날짜 TEXT NOT NULL,
    업체 TEXT NOT NULL,
    미납여부 BOOLEAN DEFAULT FALSE,
    미납내용 TEXT DEFAULT '',
    조치항목 TEXT DEFAULT '',
    수정시각 TIMESTAMPTZ DEFAULT NOW(),
    UNIQUE(날짜, 업체)
);
```

`UNIQUE` 제약이 핵심입니다. 같은 업체·같은 달(또는 같은 날짜) 데이터를 다시 올리면 새 행이 추가되지 않고 **기존 행이 갱신(upsert)** 됩니다.

### Storage 버킷 (`uploads`)

```
uploads/
├── lg/monthly/<파일명>.xls
├── lg/daily/<파일명>.csv
├── samlip/monthly/...
└── ...
```

**한글 경로는 Supabase Storage가 거부하므로** 업체명을 영문 코드로 변환해 사용합니다. 매핑은 코드의 `COMPANY_CODES` 에 정의.

### RLS (Row Level Security)
- 모든 테이블에서 **RLS는 비활성화** 상태로 운영 중.
- anon key + 앱 비밀번호로 접근 제어.
- 더 엄밀한 보안이 필요하면 RLS 정책 또는 service_role 키 + 서버 사이드 인증을 검토.

---

## 4. app.py 코드 구조

`app.py` 는 한 파일로 구성돼 있고, 위에서 아래로 다음 순서를 따릅니다.

### 4-1. 임포트와 페이지 설정

```python
import io, re, csv, os, sys, unicodedata, datetime
import pandas as pd
import altair as alt
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="판매 대시보드", page_icon="📊", layout="wide")
```

- `io`, `csv`, `unicodedata` 는 CSV 파싱과 경로 안전화에 사용
- `altair` 는 차트(막대·도넛) 라이브러리
- `set_page_config` 의 `layout="wide"` 가 양옆 여백 없는 와이드 레이아웃을 만들어줌

### 4-2. 인증 (check_password)

```python
def check_password() -> bool:
    """앱 비밀번호 체크. 세션 단위로 한 번만 묻는다."""
```

- `st.session_state["auth_ok"]` 에 인증 상태 저장 → 한 번 통과하면 그 세션 내내 통과
- 비밀번호는 `st.secrets["APP_PASSWORD"]` 에서 읽음
- 함수가 `False` 를 리턴하면 `st.stop()` 으로 이후 코드 실행 중단

### 4-3. Supabase 연결 (get_supabase)

```python
@st.cache_resource
def get_supabase():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

sb = get_supabase()
BUCKET = "uploads"
```

- `@st.cache_resource` 가 클라이언트 객체를 세션 간에 재사용하게 함 (매번 새로 만들지 않음)
- `sb` 가 이후 모든 DB·Storage 접근에 사용되는 클라이언트

### 4-4. 고정 설정 (상수)

```python
COMPANIES = ["LG생활건강", "비알코리아", "에너자이저", "라벨리", "메디카",
             "남양유업", "나사라", "삼립", "티젠"]

PALETTE = [...9색...]
COMPANY_COLORS = {업체명: 색상} 매핑

COMPANY_CODES = {업체명: 영문코드} 매핑   # Storage 경로용

SELECTED_CODES = {
    "삼립": {16개 제품코드 set}            # 삼립만 특정 제품코드 합산
}
```

**업체 추가/변경 시 손볼 곳:**
1. `COMPANIES` 에 이름 추가
2. `COMPANY_CODES` 에 영문 코드 추가
3. `PALETTE` 가 9색이라 더 늘리려면 색도 추가
4. `logo/` 폴더에 같은 이름의 PNG 추가
5. DB의 기존 행(있다면) 업데이트:
   `UPDATE monthly SET 업체 = '새이름' WHERE 업체 = '옛이름';` (daily, targets, memos에도)

### 4-5. 유틸 함수

```python
def logo_for(company: str) -> str | None:
    """logo/<업체명>.png 등을 찾아 경로 반환. 없으면 None."""
```
- 확장자는 `.png .jpg .jpeg .webp` 순으로 탐색

```python
def safe_path(s: str) -> str:
    """한글·특수문자를 [A-Za-z0-9._-]만 남기고 _로 치환."""
```
- Supabase Storage가 비ASCII 파일명을 거부하므로 필요
- 파일명 정규화에 사용

```python
def company_code(name: str) -> str:
    """업체명 → 영문 코드. COMPANY_CODES에 있으면 그걸 쓰고, 없으면 safe_path."""
```
- Storage 경로 만들 때 항상 이 함수로 변환

### 4-6. 파서

#### `parse_monthly(buf, filename, company=None) -> (월, 판매금액)`

월별 결산 엑셀(`.xls`/`.xlsx`)에서 두 값을 추출합니다.

**동작 단계:**
1. `xlrd`(.xls) 또는 `openpyxl`(.xlsx)로 헤더 없이 읽음
2. 셀을 훑어 `"YYYY년 MM월"` 패턴을 정규식으로 찾아 **월** 추출
3. `"판매금액"` 셀의 열 위치를 찾음
4. **합산 분기:**
   - `company in SELECTED_CODES`(현재는 삼립만): 해당 제품코드들의 **다음 줄**(금액 줄) 판매금액 합산
   - 그 외: `"금액 계"` 행 다음 줄의 판매금액 사용

**왜 "다음 줄"인가?** 결산서가 한 품목당 두 줄(수량 줄 + 금액 줄)이라, 코드가 적힌 줄에서 +1행 내려가야 금액이 나옴.

#### `parse_daily(buf, filename, company=None) -> (날짜, 매출액)`

일별 누적 CSV에서 두 값을 추출합니다.

**동작 단계:**
1. 파일명에서 `YYYYMMDD` 정규식으로 **날짜** 추출
2. CP949로 디코딩 (한글 매장이라 EUC-KR 계열)
3. 행 단위로 순회:
   - 컬럼 수 7 미만 → 스킵
   - `company in SELECTED_CODES`이면 `row[4]`(제품코드)가 그 set에 있는 행만 처리
   - 매출액 = `row[-7]`(판매수량) × `row[-5]`(공급단가)

**왜 "뒤에서 N번째"인가?** 제품명에 쉼표가 들어가면 CSV 열이 밀려요. 앞에서 인덱스로 잡으면 깨지지만, 뒤에서 세면 늘 같은 자리에 있어서 안전합니다.

### 4-7. DB 조회 헬퍼

```python
@st.cache_data(ttl=30)
def fetch_monthly() -> pd.DataFrame
def fetch_daily() -> pd.DataFrame
def fetch_targets() -> dict   # {업체: {월: 목표금액}}
def fetch_memos(date_key=None) -> pd.DataFrame
```

- 모두 30초 캐시 (`@st.cache_data(ttl=30)`)
- 업로드·저장 직후엔 캐시가 오래된 값을 줄 수 있으니 `invalidate_cache()` 로 강제 초기화

```python
def invalidate_cache():
    """모든 fetch_* 함수의 캐시를 비움. 데이터 변경 직후에 호출."""
    fetch_monthly.clear()
    fetch_daily.clear()
    fetch_targets.clear()
    fetch_memos.clear()
```

### 4-8. 헤더 & 데이터 로드

```python
title_col, btn_col = st.columns([6, 1], vertical_alignment="center")
# 제목 옆에 새로고침 버튼

monthly_df = fetch_monthly()
daily_df = fetch_daily()
TARGETS = fetch_targets()

overall = monthly_df.groupby("업체").sum()   # 도넛용 업체별 합계
```

페이지를 그릴 때마다 이 부분이 실행돼 최신 데이터를 가져옵니다. 캐시(30초) 덕에 DB 부하는 적어요.

### 4-9. 탭 1: 📊 대시보드

```python
tab_dash, tab_upload, tab_hist, tab_target = st.tabs([...])

with tab_dash:
    main_left, main_right = st.columns([3, 2])
```

**왼쪽 (`main_left`): 업체별 카드 3×3 그리드**

1. **카드 순위 정렬:** 각 업체의 최신 일별 매출액(`latest_daily`)을 dict로 만들고, `sorted(COMPANIES, key=...)` 로 큰 순으로 재정렬
2. **카드 내용 (각 업체):**
   - 헤더: 로고 + 순위 + 업체명
   - **일별 누적 매출액 + 달성률 (LG생활건강만):**
     - `month_key = latest_date[:7]` 로 "2026-06" 추출
     - `TARGETS[name][month_key]` 에서 그 월 목표 가져옴
     - `st.metric` 의 `delta` 에 "{누적/목표*100}% 달성" 표시
   - **월별 막대그래프:** 각 막대 위에 금액 텍스트 표시 (`mark_text(dy=-8)`)

**오른쪽 (`main_right`): 도넛 + 미납현황**

1. **도넛 차트:**
   - `overall_d = overall[overall["판매금액"] > 0]` — 음수/0 제외 (도넛 각도 계산 안 깨지게)
   - 라벨용으로 "업체명 비중%" 컬럼 생성
   - 범례용으로 업체명을 "업체 (비중%)" 로 변환 (`overall_d["원래업체"]` 에 원본 백업)
   - 색 매핑: `domain=overall_d["업체"]` (변환된 이름), `range=[COMPANY_COLORS[c] for c in overall_d["원래업체"]]` (원래 이름으로 색 조회)
   - `mark_arc` 로 도넛 본체 + `mark_text(radius=170)` 로 도넛 바깥 라벨
2. **미납현황 입력:**
   - `st.date_input` 으로 날짜 선택 → 그 날짜의 기존 메모를 DB에서 조회 (`fetch_memos(date_key)`)
   - 9개 업체 한 줄씩 표(`st.data_editor`)에 표시
   - 저장 시 **체크된 행만** upsert (체크 안 된 행은 건너뜀 — 미납 이력 보존)

### 4-10. 탭 2: 📁 파일 업로드

```python
with tab_upload:
    up_company = st.selectbox(...)
    up_kind = st.radio(...)    # 월별 / 일별
    files = st.file_uploader(..., accept_multiple_files=True)
```

**업로드 처리 흐름 (`if files and st.button(...):` 안):**

```python
for f in files:
    raw = f.getvalue()                    # 바이트 한 번 읽음
    path = f"{company_code(...)}/{kind}/{safe_path(filename)}"

    # 1) Storage에 원본 저장 (upsert로 덮어쓰기)
    sb.storage.from_(BUCKET).upload(path, raw, file_options={"upsert": "true", ...})

    # 2) 파싱
    if up_kind.startswith("월별"):
        month, sales = parse_monthly(io.BytesIO(raw), filename, up_company)
        # 3) DB에 upsert
        sb.table("monthly").upsert({...}, on_conflict="업체,월").execute()
    else:
        date, revenue = parse_daily(io.BytesIO(raw), filename, up_company)
        sb.table("daily").upsert({...}, on_conflict="업체,날짜").execute()

progress.empty()
invalidate_cache()                        # 캐시 비움 → 대시보드 즉시 반영
st.success(f"완료: 성공 N건 · 실패 M건")
```

**중요한 패턴:**
- `io.BytesIO(raw)` 로 같은 바이트를 두 번(Storage 저장 + 파싱) 사용
- `on_conflict="업체,월"` 또는 `"업체,날짜"` — UNIQUE 제약과 같은 컬럼 조합. 충돌 시 새 행 만들지 않고 업데이트
- 마지막에 `invalidate_cache()` 호출 — 안 부르면 30초간 옛 데이터가 보임

### 4-11. 탭 3: 💰 매출 이력

```python
with tab_hist:
    view = st.radio(["월별 매출", "일별 누적 매출", "미납현황"])
    # 업체·기간 필터 → DataFrame → 표 출력
```

**처리 단계:**
1. `view` 에 따라 `fetch_monthly` / `fetch_daily` / `fetch_memos` 호출
2. 업체·시작·종료 필터 적용
3. 시각 컬럼(`업로드시각`, `수정시각`) 처리:
   ```python
   kst = datetime.timezone(datetime.timedelta(hours=9))
   df["수정시각"] = pd.to_datetime(df["수정시각"], utc=True).dt.tz_convert(kst).dt.strftime(...)
   ```
   - Supabase는 UTC로 저장하므로 한국 시간(UTC+9)으로 변환해 표시
4. `st.dataframe` 에 `column_config={"판매금액": NumberColumn(format="%,d")}` 로 콤마 표시
5. CSV 다운로드 버튼: `encoding="utf-8-sig"` (Excel에서 한글 깨짐 방지)

### 4-12. 탭 4: 🎯 목표 설정

```python
with tab_target:
    company = "LG생활건강"                # 고정 (다른 업체는 목표 없음)
    sel_target_month = st.text_input("월 (YYYY-MM)", value=오늘월)
    new_target = st.number_input(...)
    
    if 저장 클릭:
        sb.table("targets").upsert({업체, 월, 목표금액}, on_conflict="업체,월").execute()
        invalidate_cache()
    
    # 저장된 목표 이력 표
    rows = sb.table("targets").select("*").eq("업체", company).order("월", desc=True).execute().data
    st.dataframe(...)
```

- 매월 새 목표를 저장하면 옛 월 목표는 자동 보존 (UNIQUE 충돌 없음 — 월이 다르니까)
- LG생활건강 외 다른 업체로 확장하려면 `selectbox` 추가 + `eq("업체", company)` 필터 조정

---

## 5. 데이터 처리 규칙

### 월별 엑셀 (.xls / .xlsx)
- "YYYY년 MM월" 패턴에서 월 추출
- 기본: "금액 계" 행의 다음 줄 → 25번 열의 판매금액
- **삼립 예외:** 16개 제품코드만 합산
  ```
  220176, 220177, 220178, 230226, 230229, 230232, 240218, 240219, 240220,
  250225, 250226, 250228, 260205, 260206, 260207, 260209
  ```

### 일별 CSV (CP949 인코딩)
- 파일명 앞 `YYYYMMDD` = 조회일 (그 달 1일~조회 전일까지의 **누적**)
- 매출액 = `판매수량 × 공급단가` (배수 없음)
- 컬럼: 판매수량 = 뒤에서 7번째, 공급단가 = 뒤에서 5번째
- 제품코드는 `row[4]` 에 있음
- 삼립은 위 16개만 합산

### Storage 경로
- 한글 불가 → 업체명을 영문 코드로 변환
- 매핑: `LG생활건강 → lg`, `비알코리아 → br`, `에너자이저 → energizer`, `라벨리 → labelly`, `메디카 → medica`, `남양유업 → namyang`, `나사라 → nasara`, `삼립 → samlip`, `티젠 → tjeen`

---

## 6. 운영 가이드

### 로컬 실행

```bash
pip install -r requirements.txt
streamlit run app.py
```

### 클라우드 배포 (Streamlit Cloud)
1. GitHub 퍼블릭 repo에 `app.py`, `requirements.txt`, `.gitignore`, `logo/` 업로드
2. https://share.streamlit.io 에서 repo 연결
3. Advanced settings → Secrets 에 `secrets.toml` 내용 붙여넣기
4. Deploy

코드를 GitHub에 commit하면 **Streamlit Cloud가 자동으로 감지해 재배포**합니다 (2~5분 소요).

### 사용량 확인
- **Supabase**: Project Settings → Usage
  - 무료 한도: DB 500MB, Storage 1GB, Egress 5GB/월
- **Streamlit Cloud**: 앱 페이지 → Analytics
  - 7일간 미접속 시 슬립 (다음 접속 시 30초~1분 깨어남)

### 데이터 삭제 (Supabase SQL Editor)

```sql
-- 일별 매출 전체 비우기
DELETE FROM daily;

-- 미납현황 전체 비우기
DELETE FROM memos;

-- 특정 업체·기간만
DELETE FROM daily WHERE 업체 = 'LG생활건강' AND 날짜 < '2026-06-01';

-- 업체명 통일 (옛 이름 → 새 이름)
UPDATE monthly SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
UPDATE daily   SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
UPDATE targets SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
UPDATE memos   SET 업체 = 'LG생활건강' WHERE 업체 = '엘지';
```

삭제 전에 Table Editor에서 **Export → CSV** 로 백업하는 습관 권장.

### 비밀번호 변경
- `secrets.toml` 의 `APP_PASSWORD` 수정 (로컬 + Streamlit Cloud 양쪽)
- 3~6개월에 한 번 변경 권장

---

## 7. 자주 막히는 지점

| 증상 | 원인 | 해결 |
|---|---|---|
| `Invalid key: <한글경로>` | Storage 경로에 한글 | `COMPANY_CODES` 매핑 확인. 새 업체 추가 시 잊지 말 것 |
| `row violates row-level security policy` | RLS가 켜져 있음 | SQL Editor: `ALTER TABLE <name> DISABLE ROW LEVEL SECURITY;` |
| `KeyError: '엘지'` 등 | DB의 업체명이 코드의 `COMPANIES` 와 다름 | `UPDATE ... SET 업체 = '...'` 로 통일 |
| 도넛이 깨져 보임 | 음수/0 합계 업체가 섞임 | `overall[overall["판매금액"] > 0]` 필터 (이미 적용됨) |
| 업로드 직후 대시보드에 안 보임 | 캐시(30초) | 🔄 새로고침 버튼 또는 `invalidate_cache()` 누락 확인 |
| 배포 후 `ModuleNotFoundError` | `requirements.txt` 누락 | 패키지 6개 다 있는지 확인 |
| `secrets.toml` 관련 에러 | Streamlit Cloud Secrets 미등록 | Manage app → Settings → Secrets |
| 같은 파일 재업로드 후 합계 이상 | `upsert` 정상 동작인데 캐시 문제 | 새로고침 |
| 시각 표시가 9시간 빠름 | UTC ↔ KST 변환 누락 | `tz_convert(kst)` 확인 |

### 인수인계 체크리스트

후임자에게 넘길 때 확인할 것:

1. **GitHub repo 접근권** — Collaborators 에 후임자 계정 추가
2. **Streamlit Cloud 접근권** — App owner 변경 또는 viewer 추가
3. **Supabase 프로젝트 접근권** — 팀원 초대
4. **`secrets.toml` 의 키 3개** — 보안 채널로 전달
5. **앱 비밀번호** — 별도 보안 채널로 전달, 가능하면 인수인계 후 변경
6. **백업 정책** — 정기적으로 monthly/daily/targets/memos를 CSV로 export하는 절차 합의
7. **운영 캘린더** — 매월 결산서 도착 시점, 업로드 일정 공유

---

## 부록: 코드 흐름 한눈에

```
앱 시작
  ↓
check_password() → False면 st.stop()
  ↓
get_supabase() → sb 클라이언트 준비
  ↓
fetch_monthly/daily/targets() → 데이터 로드 (캐시 30초)
  ↓
st.tabs([...]) → 탭 4개 생성
  ↓
[대시보드] 카드 그리드 + 도넛 + 미납현황
[업로드]   파일 받기 → 파싱 → Storage 저장 + DB upsert → invalidate_cache
[이력]     필터 → DataFrame → st.dataframe + CSV 다운로드
[목표]     입력 → targets 테이블에 upsert → invalidate_cache
```

새 기능을 추가할 때 주로 손볼 곳:
- **새 업체 추가:** `COMPANIES`, `COMPANY_CODES`, `PALETTE`, `logo/`, DB 정리
- **새 파일 형식:** `parse_monthly` 또는 `parse_daily` 안의 컬럼 위치 조정
- **새 시각화:** Altair 차트 추가 — `mark_bar`, `mark_arc`, `mark_line` 등
- **새 탭:** `st.tabs([...])` 에 항목 추가 + `with tab_xxx:` 블록 작성

문서 끝.
