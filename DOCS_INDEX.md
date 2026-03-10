# 📚 인별납부내역자동화 문서 인덱스

프로젝트 내 모든 문서 파일의 빠른 탐색을 위한 인덱스

---

## 문서 목록

### [`AGENTS.md`](AGENTS.md)

**AI 어시스턴트용 프로젝트 컨텍스트**

- 기술 스택: Python, Streamlit, Pandas, openpyxl
- 아키텍처: 계층 구조, 데이터 흐름, 오류 검출 로직
- 주요 파일: app.py, data_processor.py, constants.py
- 컨벤션: 컬럼명 규칙, 상태값, UI 스타일

| 관련 기능 | 관련 파일 |
|----------|----------|
| 오류 검출 | `data_processor.py` |
| UI 구현 | `app.py` |
| 상수 정의 | `constants.py` |

---

### [`README.md`](README.md)

**프로젝트 소개 및 빠른 시작 가이드**

- 개요 및 핵심 기능
- 설치 및 실행 방법
- 사용법 및 입력 데이터 형식

| 관련 기능 | 관련 파일 |
|----------|----------|
| 설치/실행 | `requirements.txt`, `app.py` |
| 테마 설정 | `.streamlit/config.toml` |

---

### [`PROJECT_ANALYSIS.md`](PROJECT_ANALYSIS.md)

**아키텍처 결정 사항 및 개발 가이드**

- 아키텍처 결정 사항 (스택 선택 이유, 트레이드오프)
- 데이터 흐름 원칙 (상태 관리, 캐싱)
- 핵심 비즈니스 로직 (오류 검출 규칙, 기간별 비율)
- 개발 시 주의사항 (자주 발생하는 실수, 성능, 보안)

| 관련 기능 | 관련 파일 |
|----------|----------|
| 오류 검출 규칙 | `data_processor.py` |
| 상태 관리 | `app.py` |
| 상수 참조 | `constants.py` |

---

## 기능별 문서 매핑

### 📊 데이터 처리

| 기능 | 추천 문서 |
|------|----------|
| 오류 검출 규칙 | `PROJECT_ANALYSIS.md` > 핵심 비즈니스 로직 |
| 최초납부월 필터링 | `PROJECT_ANALYSIS.md` > 핵심 비즈니스 로직 |
| 엑셀 내보내기 | `AGENTS.md` > 핵심 데이터 흐름 |

### 🎨 UI/UX

| 기능 | 추천 문서 |
|------|----------|
| Streamlit 구조 | `AGENTS.md` > Architecture |
| 테마 설정 | `.streamlit/config.toml` 직접 확인 |
| 노르딕 브루탈리즘 | `AGENTS.md` > UI 스타일 |

### 🔧 개발

| 기능 | 추천 문서 |
|------|----------|
| 시작하기 | `README.md` > 시작하기 |
| 컬럼명 규칙 | `AGENTS.md` > Conventions |
| 자주 발생하는 실수 | `PROJECT_ANALYSIS.md` > 개발 시 주의사항 |
| 성능/보안 | `PROJECT_ANALYSIS.md` > 개발 시 주의사항 |


