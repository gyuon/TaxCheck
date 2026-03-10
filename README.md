# 인별납부내역자동화 (TaxCheck)

졸업생 기금 납부 내역 분석 및 오류 자동 검출 도구

## 개요

Streamlit 기반 웹 애플리케이션으로, 졸업생의 기금 납부 내역을 분석하여 다음 오류를 자동으로 검출합니다:

- **미납**: 해당 월에 납부 기록이 없는 경우
- **부족**: 기준금액보다 적게 납부한 경우
- **초과**: 기준금액보다 많이 납부한 경우

## 기술 스택

- Python 3.12
- Streamlit 1.34+
- Pandas 2.0+
- openpyxl 3.1+
- NumPy 1.26+

## 시작하기

### 사전 요구사항

- Python 3.12+
- pip

### 설치 방법

```bash
# 저장소 클론
cd 인별납부내역자동화

# 가상환경 생성 및 활성화
python -m venv TaxCheck/venv
source TaxCheck/venv/bin/activate  # Linux/Mac
# TaxCheck\venv\Scripts\activate   # Windows

# 의존성 설치
pip install -r TaxCheck/requirements.txt
```

### 실행 방법

```bash
cd TaxCheck
streamlit run app.py
```

브라우저에서 `http://localhost:8501` 접속

## 환경 변수

현재 환경 변수 설정 불필요

## 사용법

1. **원본 엑셀 업로드**: 납부 내역이 포함된 엑셀 파일 (.xlsx)
2. **졸업생 명단 연결**: Google Sheets URL 입력
3. **분석 기간 설정**: 시작년도 ~ 종료년도
4. **분석 실행**: 오류 검출 결과 확인 및 엑셀 다운로드

## 입력 데이터 형식

### 원본 엑셀 (필수 컬럼)
- `이름`: 납부자명
- `코드1`: 기금 구분 (1=운영, 2=협력, 3=복지)
- `코드2`: 세부 코드
- `입금`: 납부 금액
- `해당년`/`해당월`: 납부 년월

### Google Sheets (졸업생 명단)
- `이름`: 졸업생명
- `구분`: "졸업생" 값이 있는 행만 필터링

## 라이선스

MIT License
