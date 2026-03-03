# Changelog

모든 주요 변경사항은 이 파일에 기록됩니다.

형식은 [Keep a Changelog](https://keepachangelog.com/ko/1.0.0/)를 기반으로 하며,
[Semantic Versioning](https://semver.org/lang/ko/)을 따릅니다.

## [Unreleased]

## [1.0.0] - 2025-02-25

### Added
- Streamlit 기반 웹 애플리케이션 초기 구현
- 엑셀 파일 업로드 및 분석 기능
- Google Sheets 연동 (졸업생 명단 필터링)
- 기금별 오류 검출 로직
  - 운영기금 미납 검출 (코드 11~17)
  - 협력기금 부족/초과 검출 (운영기금 30%~40%)
  - 복지기금 부족/초과 검출 (운영기금 100%)
- 최초 납부월 추출 및 미납 필터링
- 분석 결과 엑셀 내보내기 (3개 시트)
  - 오류검출결과
  - 오류요약 (이름별)
  - 최초납부월
- 노르딕 브루탈리즘 UI 테마 적용
- 단위 테스트 스위트 구현

### Changed
- 컬럼명 변경: `해당년` → `납부년`, `해당월` → `납부월`
- 협력기금 비율: 2019년 3월 이전 30%, 이후 40% 적용
- FutureWarning 해결 (pandas 다운캐스팅, dtype 정렬)

### Technical
- pandas 2.0+ 호환성 확보
- openpyxl 스타일링 최적화
- Streamlit 세션 상태 관리 구현
