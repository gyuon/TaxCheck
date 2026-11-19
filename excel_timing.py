import logging
import sys
import time

try:
    import resource
except ImportError:
    resource = None


log = logging.getLogger("taxcheck")


def _peak_memory_mb():
    if resource is None:
        return "unavailable"
    try:
        peak = resource.getrusage(resource.RUSAGE_SELF).ru_maxrss
        return round(peak / (1024 * 1024 if sys.platform == "darwin" else 1024), 1)
    except OSError:
        return "unavailable"


def start_excel_stage(stage, **details):
    """종료 함수를 반환한다. CPU는 현재 스레드, 메모리는 프로세스 누적 피크다."""
    started = time.perf_counter()
    cpu_started = time.thread_time()
    log.info("excel_stage event=start stage=%s details=%s", stage, details)

    def finish(**result):
        wall = time.perf_counter() - started
        cpu = time.thread_time() - cpu_started
        log.info(
            "excel_stage event=complete stage=%s wall_s=%.3f cpu_s=%.3f peak_rss_mb=%s details=%s",
            stage, wall, cpu, _peak_memory_mb(), {**details, **result},
        )

    return finish
