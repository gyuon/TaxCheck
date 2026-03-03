# 📚 인별납부내역자동화 문서 인덱스

프로젝트 내 모든 문서 파일의 빠른 탐색을 위한 인덱스

---

## 문서 목록

### [`AGENTS.md`](AGENTS.md)

**AI 어시스턴트용 프로젝트 컨텍스트**

- 기술 스택: Python, Streamlit, Pandas, openpyxl
- 아키텍처: 디렉토리 구조, 데이터 흐름, 오류 검출 로직
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
- 프로젝트 구조
- 사용법 및 입력 데이터 형식

| 관련 기능 | 관련 파일 |
|----------|----------|
| 설치/실행 | `requirements.txt`, `app.py` |
| 테마 설정 | `.streamlit/config.toml` |

---

### [`CHANGELOG.md`](CHANGELOG.md)

**버전별 변경 이력**

- v1.0.0 (2025-02-25): 초기 릴리스
- 기능 추가, 변경사항, 기술적 개선

| 관련 기능 | 관련 파일 |
|----------|----------|
| 릴리스 관리 | 모든 프로젝트 파일 |

---

### [`PROJECT_ANALYSIS.md`](PROJECT_ANALYSIS.md)

**전체 프로젝트 심층 분석**

- 데이터 스키마 (입력/출력)
- 핵심 기능 플로우 다이어그램
- 오류 검출 로직 상세 설명
- 주요 함수 구조
- 보안/성능 고려사항

| 관련 기능 | 관련 파일 |
|----------|----------|
| 데이터 처리 | `data_processor.py` |
| UI 흐름 | `app.py` |
| 테스트 | `test_suite.py` |

---

## 기능별 문서 매핑

### 📊 데이터 처리

| 기능 | 추천 문서 |
|------|----------|
| 오류 검출 로직 | `PROJECT_ANALYSIS.md` > 핵심 기능 플로우 |
| 미납월 생성 | `PROJECT_ANALYSIS.md` > 오류 검출 로직 상세 |
| 엑셀 내보내기 | `AGENTS.md` > 핵심 데이터 흐름 |

### 🎨 UI/UX

| 기능 | 추천 문서 |
|------|----------|
| Streamlit 구조 | `AGENTS.md` > Architecture |
| 테마 설정 | `README.md` > 프로젝트 구조 |
| 노르딕 브루탈리즘 | `AGENTS.md` > UI 스타일 |

### 🔧 개발

| 기능 | 추천 문서 |
|------|----------|
| 시작하기 | `README.md` > 시작하기 |
| 컬럼명 규칙 | `AGENTS.md` > Conventions |
| 테스트 작성 | `PROJECT_ANALYSIS.md` > 개발 참고사항 |

---

## 문서 업데이트 이력

| 날짜 | 문서 | 변경 내용 |
|------|------|----------|
| 2025-03-04 | `DOCS_INDEX.md` | 파일 경로 참조 수정 (TaxCheck/ 접두사 제거) |
| 2025-03-04 | `AGENTS.md` | Common Commands 섹션 경로 수정 |
| 2025-03-04 | `PROJECT_ANALYSIS.md` | 개발 참고사항 섹션 경로 수정 |
| 2025-03-04 | `DOCS_INDEX.md` | 초기 생성 |
| 2025-03-04 | `AGENTS.md` | 초기 생성 |
| 2025-03-04 | `README.md` | 초기 생성 |
| 2025-03-04 | `CHANGELOG.md` | v1.0.0 릴리스 기록 |
| 2025-03-04 | `PROJECT_ANALYSIS.md` | 초기 생성 |
