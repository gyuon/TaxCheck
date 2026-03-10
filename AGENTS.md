# AGENTS.md

## Project Overview

**프로젝트명**: 인별납부내역자동화 (TaxCheck)

**목적**: 졸업생 기금(운영기금/협력기금/복지기금) 납부 내역을 분석하여 오류(미납, 부족, 초과)를 자동 검출하는 Streamlit 웹 애플리케이션

**타겟 사용자**: 기금 관리 담당자, 행정 실무자

## Tech Stack

| 계층 | 기술 |
|------|------|
| Frontend | Streamlit 1.34+, st-aggrid |
| Backend | Python 3.12 |
| Data Processing | Pandas 2.0+, NumPy 1.26+ |
| Excel | openpyxl 3.1+ |
| External API | Google Sheets (CSV export), requests |

## Common Commands

```bash
# 가상환경 활성화
source venv/bin/activate

# 의존성 설치
pip install -r requirements.txt

# Streamlit 앱 실행
streamlit run app.py

# 테스트 실행
python -m unittest test_suite.py
```

## Environment Variables

현재 환경 변수 없음 (Google Sheets는 공개 URL 사용)

## Architecture

**계층 구조**: UI 계층(app.py)과 비즈니스 로직 계층(data_processor.py)이 분리된 단일 모듈 구조. 상수는 constants.py에서 중앙 관리.

**상태 관리**: Streamlit 세션 상태(st.session_state)로 파일 업로드, 처리 결과, UI 상태를 관리.

**데이터 처리 패턴**: 모든 데이터는 Pandas DataFrame으로 메모리에서 처리되며, DB는 사용하지 않음.

### 핵심 데이터 흐름

1. **입력**: 엑셀 파일(원본 납부내역) + Google Sheets URL(졸업생 명단)
2. **처리**:
   - 졸업생 명단으로 필터링
   - 최초 납부월 추출
   - 년도 범위 필터링
   - 오류 검출 (detect_errors)
   - 미납월 생성 (generate_missed_months)
3. **출력**: 엑셀 파일 (오류검출결과, 오류요약, 최초납부월 시트)

### 오류 검출 로직

| 기금 | 기준 | 오류 유형 |
|------|------|-----------|
| 운영기금 | 코드 11~17 중 하나라도 납부 | 미납 (모두 없음) |
| 협력기금 | 운영기금 × 30% (2019.03 이전) / 40% (이후) | 부족/초과 |
| 복지기금 | 운영기금 × 100% | 부족/초과 |

## Key Files

| 파일 | 역할 |
|------|------|
| `app.py` | Streamlit UI, 세션 상태 관리 |
| `data_processor.py` | 오류 검출, 미납 생성 등 비즈니스 로직 |
| `constants.py` | 컬럼명, 상태값, 기금 코드 상수 |

## Conventions

### 컬럼명 규칙
- 원본 데이터: `이름`, `코드1`, `코드2`, `입금`, `해당년`, `해당월`
- 결과 데이터: `이름`, `납부년`, `납부월`, `기금명`, `코드`, `납부금액`, `기준금액`, `상태`, `차액`, `비고`

### 상태값
- `미납` (UNPAID): 해당 월에 납부 기록 없음
- `부족` (INSUFFICIENT): 기준금액보다 1,000원 이상 적게 납부
- `초과` (EXCESS): 기준금액보다 1,000원 이상 많이 납부

### UI 스타일
- 노르딕 브루탈리즘 디자인 (고대비, 직각 모서리, 굵은 테두리)
- 주 색상: #1F2937 (Dark Gray), #F2C94C (Sunflower Yellow)

## Known Issues / TODO

- `compare_excel.py`의 Windows 경로 하드코딩 제거 필요
- Google Sheets API 인증 추가 고려 (현재 공개 URL만 지원)
