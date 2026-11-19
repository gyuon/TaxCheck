import logging
import platform
import sys
import time
from contextvars import ContextVar
from functools import lru_cache, wraps
from importlib.metadata import PackageNotFoundError, version
from pathlib import Path
from uuid import uuid4

try:
    import resource
except ImportError:
    resource = None


log = logging.getLogger("taxcheck")
_run_id = ContextVar("excel_run_id", default="untracked")


def new_excel_run_id():
    return uuid4().hex


def log_excel_event(event, run_id=None, **details):
    log.info("excel_diagnostic event=%s run_id=%s details=%s", event, run_id or _run_id.get(), details)


@lru_cache(maxsize=1)
def _runtime_environment():
    from openpyxl.xml import LXML

    environment = {"python": platform.python_version(), "lxml_enabled": LXML}
    for package in ("pandas", "openpyxl", "lxml"):
        try:
            environment[package] = version(package)
        except PackageNotFoundError:
            environment[package] = "not_installed"
        except (OSError, ValueError):
            environment[package] = "unavailable"
    return environment


def _read_text(path):
    try:
        return Path(path).read_text()
    except (OSError, UnicodeError):
        return ""


def _cpu_cgroup():
    """가시적인 Linux cgroup의 경로를 찾는다. 제한이 상위 계층에 있을 수도 있다."""
    root = Path("/sys/fs/cgroup")
    for line in _read_text("/proc/self/cgroup").splitlines():
        fields = line.split(":", 2)
        if len(fields) != 3:
            continue
        _, controllers, relative = fields
        # cgroup namespace에 따라 /sys/fs/cgroup 자체가 현재 그룹일 수 있다.
        if controllers == "":
            for path in (root / relative.lstrip("/"), root):
                if (path / "cpu.stat").is_file():
                    return 2, path
        elif "cpu" in controllers.split(","):
            for mount in ("cpu,cpuacct", "cpu"):
                for path in (root / mount / relative.lstrip("/"), root / mount):
                    if (path / "cpu.stat").is_file():
                        return 1, path
    return None, None


def _cpu_snapshot():
    try:
        cgroup_version, path = _cpu_cgroup()
        if path is None:
            return {"available": False}
        stats = dict(line.split() for line in _read_text(path / "cpu.stat").splitlines())
        if not stats:
            return {"available": False}
        result = {"available": True, "version": cgroup_version}
        for key in ("nr_periods", "nr_throttled"):
            if key in stats:
                result[key] = int(stats[key])
        if cgroup_version == 2:
            for key in ("usage_usec", "throttled_usec"):
                if key in stats:
                    result[key] = int(stats[key])
        else:
            if "throttled_time" in stats:
                result["throttled_usec"] = int(stats["throttled_time"]) // 1000
        try:
            if cgroup_version == 2:
                quota, period = _read_text(path / "cpu.max").split()
                result["cpu_limit_cores"] = "unlimited" if quota == "max" else round(int(quota) / int(period), 4)
            else:
                quota = int(_read_text(path / "cpu.cfs_quota_us"))
                period = int(_read_text(path / "cpu.cfs_period_us"))
                result["cpu_limit_cores"] = "unlimited" if quota < 0 else round(quota / period, 4)
        except (ValueError, ZeroDivisionError):
            result["cpu_limit_cores"] = "unavailable"
        return result
    except (OSError, ValueError, ZeroDivisionError):
        return {"available": False}


def _cpu_delta(before, after):
    if not before.get("available") or not after.get("available") or before.get("version") != after.get("version"):
        return {"available": False}
    return {
        key: after[key] - before[key]
        for key in ("usage_usec", "nr_periods", "nr_throttled", "throttled_usec")
        if key in before and key in after and after[key] >= before[key]
    }


def _current_memory_mb():
    for line in _read_text("/proc/self/status").splitlines():
        if line.startswith("VmRSS:"):
            try:
                return round(int(line.split()[1]) / 1024, 1)
            except (ValueError, IndexError):
                break
    return "unavailable"


def trace_excel_export(callback):
    """스레드 안에서 ID를 설정한다. 동시 요청 간 ContextVar를 공유하지 않는다."""
    @wraps(callback)
    def traced(*args, diagnostic_run_id=None, diagnostic_submitted_at=None, **kwargs):
        token = _run_id.set(diagnostic_run_id or new_excel_run_id())
        started, thread_started, process_started = time.perf_counter(), time.thread_time(), time.process_time()
        before = {"available": False}
        outcome = "worker_failed"
        try:
            before = _cpu_snapshot()
            log_excel_event(
                "worker_start", queue_wait_s=None if diagnostic_submitted_at is None else round(started - diagnostic_submitted_at, 3),
                runtime=_runtime_environment(), cgroup_before=before,
                current_rss_mb=_current_memory_mb(),
            )
            result = callback(*args, **kwargs)
            outcome = "worker_complete"
            return result
        finally:
            try:
                after = _cpu_snapshot()
                log_excel_event(
                    outcome, wall_s=round(time.perf_counter() - started, 3),
                    thread_cpu_s=round(time.thread_time() - thread_started, 3),
                    process_cpu_s=round(time.process_time() - process_started, 3),
                    cgroup_after=after, cgroup_delta=_cpu_delta(before, after),
                    current_rss_mb=_current_memory_mb(), peak_rss_mb=_peak_memory_mb(),
                )
            finally:
                _run_id.reset(token)
    return traced


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
    log.info("excel_stage event=start stage=%s run_id=%s details=%s", stage, _run_id.get(), details)

    def finish(**result):
        wall = time.perf_counter() - started
        cpu = time.thread_time() - cpu_started
        log.info(
            "excel_stage event=complete stage=%s wall_s=%.3f cpu_s=%.3f peak_rss_mb=%s run_id=%s details=%s",
            stage, wall, cpu, _peak_memory_mb(), _run_id.get(), {**details, **result},
        )

    return finish
