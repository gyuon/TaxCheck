# 결과 파일 생성 성능 로그

Render Dashboard에서 해당 서비스의 **Logs**를 열고 `excel_`로 검색한다.
같은 작업의 `run_id`를 복사해 검색하면 동시 실행 로그를 구분할 수 있다.
이름, 원본 파일명, 납부 금액 등 개인 데이터는 진단 로그에 기록하지 않는다.

## 작업 흐름

- `excel_diagnostic event=submitted`: 서버에서 파일 생성을 요청한 시점.
- `worker_start`: 작업 스레드 시작. `queue_wait_s`는 요청부터 스레드 시작까지의 대기 시간이다.
  `runtime`에는 Python/pandas/openpyxl/lxml 버전과 `lxml_enabled`가 포함된다.
  lxml은 `requirements.txt`의 배포 의존성이다. 배포 후 `lxml_enabled: True`와
  설치 버전을 확인한다. `OPENPYXL_LXML=False` 환경 변수는 이 가속을 비활성화한다.
- `excel_stage`: 기존 단계별 시작/완료 로그. `wall_s`는 실제 경과 시간,
  `cpu_s`는 현재 작업 스레드의 CPU 사용 시간이다. 부모 단계와 하위 단계를 합산하면 중복된다.
- `save_breakdown`: `workbook_save_zip` 안의 저장 시간 구분.
  `zip_write_*`는 ZIP 항목 쓰기와 닫기 시간으로, 압축·임시 XML 파일 읽기·아카이브 쓰기가 포함된다.
  `xml_and_other_*`는 이를 제외한 나머지 저장 시간으로, XML 직렬화와 기타 저장 준비를 포함한다.
  순수 XML/압축 시간으로 해석하지 않는다. `uncompressed_bytes`/`compressed_bytes`는 ZIP 항목 크기 합계이며
  ZIP 헤더를 포함한 최종 파일 크기와 다르다.
- `worker_complete` 또는 `worker_failed`: 스레드 작업 종료 및 자원 통계.
- `결과 엑셀 생성 완료`: 서버 코루틴이 결과를 받은 시점. 스레드 대기와 반환 대기도 포함한다.
- `ready_update_enqueued`: 다운로드 버튼 변경을 서버에서 전송 대기열에 넣은 시점.
  실제 브라우저 표시 완료를 뜻하지 않는다. `discarded`는 무효화된 이전 작업의 결과다.

## 자원 통계 해석

`worker_start`의 `cgroup_before`와 종료 시 `cgroup_after`는 Linux에서 보이는 CPU cgroup 통계다.
지원하지 않는 운영체제나 읽을 수 없는 환경에서는 `available: False`로 표시한다.

- `cpu_limit_cores`: 확인된 그룹의 CPU quota / period. `0.1`이면 CPU 0.1개 분량이다.
  `unlimited`여도 상위 그룹이나 호스트에서 제한될 수 있다.
- `cgroup_delta.nr_throttled`: 작업 구간 중 CPU quota가 소진된 period 수 증가분.
- `cgroup_delta.throttled_usec`: CPU 제한으로 실행이 지연된 시간 증가분(마이크로초).
- `cgroup_delta.usage_usec`: cgroup v2의 CPU 사용 시간 증가분.
- 이 수치들은 해당 cgroup 전체를 나타내며 동시에 실행된 다른 작업도 포함한다.
  제한 시간은 여러 CPU에서 합산될 수 있으므로 작업 지연 시간과 일대일 대응시키지 않는다.
- `thread_cpu_s`는 작업 스레드, `process_cpu_s`는 프로세스 전체의 CPU 사용 시간이다.
- `current_rss_mb`는 측정 시점의 프로세스 RSS, `peak_rss_mb`는 프로세스 시작 이후의 누적 최대 RSS다.

같은 입력 파일과 분석 기간으로 테스트하고 `submitted`부터 `ready_update_enqueued`까지 공유하면 된다.
서버 완료 이후에도 버튼 표시가 늦으면 Chrome 개발자 도구의 Network → WS → Messages에서
수신 시각을 확인하여 서버 생성과 브라우저 반영 지연을 구분한다.
