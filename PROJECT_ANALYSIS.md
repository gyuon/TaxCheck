# 인별납부내역자동화 프로젝트 분석

## 프로젝트 개요

### 목적
졸업생 기금(운영기금/협력기금/복지기금) 납부 내역을 분석하여 오류(미납, 부족, 초과)를 자동 검출하는 Streamlit 웹 애플리케이션

### 핵심 기능
1. 엑셀 파일 업로드 및 원본 납부내역 분석
2. Google Sheets 연동 졸업생 명단 필터링
3. 기금별 납부 오류 검출 (미납, 부족, 초과)
4. 최초 납부월 기준 미납 데이터 필터링
5. 분석 결과 엑셀 내보내기

## 기술 스택

| 계층 | 기술 | 버전 |
|------|------|------|
| Frontend | Streamlit | 1.34+ |
| Data Grid | st-aggrid | 0.3.4+ |
| Backend | Python | 3.12 |
| Data Processing | Pandas | 2.0+ |
| Numerical | NumPy | 1.26+ |
| Excel | openpyxl | 3.1+ |
| HTTP | requests | 2.31+ |

## 프로젝트 구조

```
인별납부내역자동화/
├── TaxCheck/
│   ├── app.py                 # Streamlit 메인 애플리케이션 (679 lines)
│   ├── data_processor.py      # 데이터 처리 로직 (664 lines)
│   ├── constants.py           # 상수 정의 (63 lines)
│   ├── test_suite.py          # 단위 테스트 (132 lines)
│   ├── compare_excel.py       # 엑셀 비교 유틸리티 (96 lines)
│   ├── requirements.txt       # 의존성 목록
│   ├── .gitignore
│   └── .streamlit/
│       └── config.toml        # 테마 설정
└── venv/                      # Python 가상환경
```

## 데이터베이스 스키마

이 프로젝트는 데이터베이스를 사용하지 않음. 모든 데이터는 메모리(Pandas DataFrame)와 엑셀 파일로 처리.

### 입력 데이터 스키마

**원본 엑셀 (raw sheet)**
| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| 이름 | string | 납부자명 |
| 코드1 | int | 기금 구분 (1=운영, 2=협력, 3=복지) |
| 코드2 | int | 세부 코드 (1~7) |
| 코드3 | int | (선택) 추가 코드 |
| 입금 | int | 납부 금액 |
| 해당년 | int | 납부 연도 |
| 해당월 | int | 납부 월 |
| 구분 | string | (선택) 구분 정보 |

**Google Sheets (졸업생 명단)**
| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| 이름 | string | 졸업생명 |
| 구분 | string | "졸업생" 여부 |

### 출력 데이터 스키마

**오류검출결과 시트**
| 컬럼명 | 타입 | 설명 |
|--------|------|------|
| 번호 | int | 순번 |
| 이름 | string | 납부자명 |
| 납부년 | int | 납부 연도 |
| 납부월 | int | 납부 월 |
| 기금명 | string | 운영/협력/복지 |
| 코드 | string | 기금 코드 |
| 납부금액 | int | 실제 납부액 |
| 기준금액 | int | 기준 납부액 |
| 기준금액 산출법 | string | 계산 공식 |
| 상태 | string | 미납/부족/초과 |
| 차액 | int | 납부액 - 기준액 |
| 비고 | string | (선택) 비고 |

## 핵심 기능 플로우

### 1. 데이터 처리 흐름

```
[엑셀 업로드] → [Google Sheets 연동] → [졸업생 필터링]
      ↓
[최초납부월 추출] → [년도 범위 필터링]
      ↓
[오류 검출] + [미납월 생성]
      ↓
[결과 병합] → [엑셀 내보내기]
```

### 2. 오류 검출 로직 상세

#### 운영기금 (코드 11~17)
- **미납 조건**: 해당 (이름, 년, 월) 그룹에 코드 11~17 중 어느 것도 없음
- **코드 표기**: 미납 시 "11~17"로 표기

#### 협력기금 (코드 21)
- **기준**: 운영기금 총액 × 비율
  - 2019년 3월 이전: 30%
  - 2019년 4월 이후: 40%
- **부족**: 기준금액보다 1,000원 이상 적게 납부
- **초과**: 기준금액보다 1,000원 이상 많이 납부

#### 복지기금 (코드 31)
- **기준**: 운영기금 총액 × 100%
- **부족/초과**: 협력기금과 동일한 기준 (1,000원 이상 차이)

### 3. 최초 납부월 필터링

```python
# 최초 납부월 판정 조건
condition = (
    ((코드1 == 1) & (코드2 in 1~7)) |
    ((코드1 in [2, 3]) & (코드2 == 1))
)
```

- 최초 납부월 이전의 미납 데이터는 결과에서 제외
- 제외된 건수는 UI에 표시

## 주요 함수 구조

### data_processor.py

```python
# 데이터 정규화
normalize_names(df) -> DataFrame

# Google Sheets 제목 추출
get_google_sheets_title(url) -> str | None

# 최초 납부월 추출
extract_first_payment_month(df) -> DataFrame

# 오류 검출 (부족/초과)
detect_errors(df) -> DataFrame

# 미납월 생성
generate_missed_months(df, df_first) -> (DataFrame, int)

# 엑셀 바이트 변환
to_excel_bytes(df_first, df_errors, filename, df_summary) -> (bytes, str)
```

### app.py

```python
# Streamlit 캐시 래퍼
build_excel_bytes(...) -> (bytes, str)

# HTML 테이블 렌더링
render_styled_table(df) -> None

# 앱 상태 초기화
reset_app() -> None

# 메인 처리 로직
run_processing(main_file, gsheet_url, sheet_name, 
               header_row, start_year, end_year, use_filter) -> None
```

## 보안/성능 고려사항

### 보안
- Google Sheets는 공개 URL만 지원 (인증 없음)
- 파일 업로드 최대 크기: 20MB (Streamlit 설정)
- 민감 데이터 로그 없음

### 성능
- Pandas 벡터화 연산 사용
- Streamlit `@st.cache_data`로 엑셀 생성 결과 캐싱
- UI 렌더링 시 상위 200건만 표시 (전체는 엑셀 다운로드)
- openpyxl 스타일링 최적화

## 개발 참고사항

### 컬럼명 일관성
- `constants.py`의 `Col` 클래스에서 모든 컬럼명 중앙 관리
- `app.py`와 `data_processor.py`는 반드시 `Col` 클래스 참조

### 상태값
- `Status` 클래스에서 관리: `UNPAID`, `INSUFFICIENT`, `EXCESS`
- `app.py`의 `counts` dict와 일치해야 함

### 테스트
- `test_suite.py`로 컬럼명 일관성 및 로직 검증
- 실행: `python -m unittest test_suite.py`

### UI 테마
- `.streamlit/config.toml`에서 테마 설정
- 노르딕 브루탈리즘 스타일 CSS는 `app.py`에 인라인
