from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from threading import Barrier
import unittest
from unittest.mock import patch
from zipfile import ZipFile

import pandas as pd
from openpyxl.comments import Comment

from excel_save import DiagnosticExcelWriter
import excel_timing as timing


class TestExportDiagnostics(unittest.TestCase):
    def test_queue_wait_and_cpu_scopes_are_recorded(self):
        @timing.trace_excel_export
        def export(value):
            return value

        snapshots = [
            {"available": True, "version": 2, "nr_throttled": 2, "throttled_usec": 100},
            {"available": True, "version": 2, "nr_throttled": 5, "throttled_usec": 900},
        ]
        with patch.object(timing, "_cpu_snapshot", side_effect=snapshots), \
                patch.object(timing, "_runtime_environment", return_value={"lxml_enabled": True}), \
                patch.object(timing.time, "perf_counter", side_effect=[10, 12]), \
                patch.object(timing.time, "thread_time", side_effect=[3, 4]), \
                patch.object(timing.time, "process_time", side_effect=[5, 7]), \
                self.assertLogs("taxcheck", level="INFO") as logs:
            self.assertEqual(export("result", diagnostic_run_id="job-1", diagnostic_submitted_at=8), "result")
        start, end = [record.args[-1] for record in logs.records]
        self.assertEqual(start["queue_wait_s"], 2)
        self.assertEqual(end["wall_s"], 2)
        self.assertEqual(end["thread_cpu_s"], 1)
        self.assertEqual(end["process_cpu_s"], 2)
        self.assertEqual(end["cgroup_delta"], {"nr_throttled": 3, "throttled_usec": 800})

    def test_concurrent_exports_keep_their_own_ids(self):
        barrier = Barrier(2)

        @timing.trace_excel_export
        def export():
            finish = timing.start_excel_stage("concurrent")
            barrier.wait(timeout=10)
            finish()
            return timing._run_id.get()

        with self.assertLogs("taxcheck", level="INFO") as logs, ThreadPoolExecutor(max_workers=2) as pool:
            futures = [pool.submit(export, diagnostic_run_id=key) for key in ("first", "second")]
            self.assertEqual([future.result(timeout=10) for future in futures], ["first", "second"])
        for key in ("first", "second"):
            self.assertEqual(sum(f"run_id={key} " in message for message in logs.output), 4)
        self.assertEqual(timing._run_id.get(), "untracked")

    def test_failure_is_logged_and_context_is_restored(self):
        @timing.trace_excel_export
        def export():
            raise ValueError("original error")

        with self.assertLogs("taxcheck", level="INFO") as logs:
            with self.assertRaisesRegex(ValueError, "original error"):
                export(diagnostic_run_id="failed")
        self.assertIn("event=worker_failed run_id=failed ", logs.output[-1])
        self.assertFalse(any("worker_complete" in message for message in logs.output))
        self.assertEqual(timing._run_id.get(), "untracked")

    def test_diagnostics_work_without_linux_resources(self):
        with patch.object(timing, "_read_text", return_value=""), \
                patch.object(timing, "resource", None), \
                self.assertLogs("taxcheck", level="INFO") as logs:
            result = timing.trace_excel_export(lambda: b"excel")()
        self.assertEqual(result, b"excel")
        details = logs.records[-1].args[-1]
        self.assertEqual(details["cgroup_delta"], {"available": False})
        self.assertEqual(details["current_rss_mb"], "unavailable")
        self.assertEqual(details["peak_rss_mb"], "unavailable")

    def test_cgroup_v2_limits_and_throttling_are_read(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "cpu.stat").write_text("usage_usec 1000\nnr_periods 10\nnr_throttled 8\nthrottled_usec 4000\n")
            (root / "cpu.max").write_text("10000 100000\n")
            with patch.object(timing, "_cpu_cgroup", return_value=(2, root)):
                snapshot = timing._cpu_snapshot()
                self.assertEqual(snapshot["cpu_limit_cores"], 0.1)
                self.assertEqual(snapshot["throttled_usec"], 4000)
                (root / "cpu.max").write_text("max 100000\n")
                self.assertEqual(timing._cpu_snapshot()["cpu_limit_cores"], "unlimited")
                (root / "cpu.max").unlink()
                partial = timing._cpu_snapshot()
                self.assertEqual(partial["cpu_limit_cores"], "unavailable")
                self.assertEqual(partial["throttled_usec"], 4000)
                (root / "cpu.stat").write_text("unexpected invalid content\n")
                self.assertEqual(timing._cpu_snapshot(), {"available": False})

    def test_cgroup_v1_nanoseconds_are_converted_to_microseconds(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "cpu.stat").write_text("nr_periods 10\nnr_throttled 8\nthrottled_time 4000000\n")
            (root / "cpu.cfs_quota_us").write_text("10000\n")
            (root / "cpu.cfs_period_us").write_text("100000\n")
            with patch.object(timing, "_cpu_cgroup", return_value=(1, root)):
                snapshot = timing._cpu_snapshot()
            self.assertEqual(snapshot["cpu_limit_cores"], 0.1)
            self.assertEqual(snapshot["throttled_usec"], 4000)

    def test_cgroup_namespace_root_is_used_when_membership_path_is_not_mounted(self):
        with TemporaryDirectory() as directory:
            root = Path(directory)
            (root / "cpu.stat").write_text("nr_throttled 0\n")
            with patch.object(timing, "Path", return_value=root), \
                    patch.object(timing, "_read_text", return_value="0::/container/not/mounted\n"):
                self.assertEqual(timing._cpu_cgroup(), (2, root))

    def test_reset_cpu_counters_are_not_reported_as_negative_deltas(self):
        before = {"available": True, "version": 2, "throttled_usec": 1000}
        after = {"available": True, "version": 2, "throttled_usec": 100}
        self.assertNotIn("throttled_usec", timing._cpu_delta(before, after))

    def test_runtime_environment_records_missing_optional_lxml(self):
        def installed_version(package):
            if package == "lxml":
                raise timing.PackageNotFoundError(package)
            return "test-version"

        with patch.object(timing, "version", side_effect=installed_version):
            environment = timing._runtime_environment.__wrapped__()
        self.assertEqual(environment["lxml"], "not_installed")
        self.assertIsInstance(environment["lxml_enabled"], bool)


class TestSaveDiagnostics(unittest.TestCase):
    def _workbook(self, writer_class):
        buffer = BytesIO()
        with writer_class(buffer) as writer:
            pd.DataFrame({"이름": ["A", "B"], "금액": [1000, 2500], "날짜": [datetime(2026, 1, 1)] * 2}).to_excel(
                writer, sheet_name="test", index=False,
            )
            sheet = writer.sheets["test"]
            sheet["D2"] = "=B2+100"
            sheet["D1"].comment = Comment("memo", "TaxCheck")
            sheet.merge_cells("A5:B5")
            sheet["A5"] = "merged"
            sheet.row_dimensions[3].hidden = True
            sheet.auto_filter.ref = "A1:D3"
            sheet.freeze_panes = "A2"
        return buffer.getvalue()

    def test_measured_save_preserves_all_excel_parts(self):
        standard = self._workbook(pd.io.excel._openpyxl.OpenpyxlWriter)
        with self.assertLogs("taxcheck", level="INFO") as logs:
            measured = self._workbook(DiagnosticExcelWriter)
        with ZipFile(BytesIO(standard)) as before, ZipFile(BytesIO(measured)) as after:
            self.assertEqual(before.namelist(), after.namelist())
            for name in before.namelist():
                if name != "docProps/core.xml":  # 저장 시간만 다를 수 있다.
                    self.assertEqual(before.read(name), after.read(name), name)
            details = logs.records[-1].args[-1]
            self.assertEqual(details["zip_write_calls"], len(after.namelist()))
            self.assertEqual(details["uncompressed_bytes"], sum(info.file_size for info in after.infolist()))
            self.assertEqual(details["compressed_bytes"], sum(info.compress_size for info in after.infolist()))
        for clock in ("wall", "cpu"):
            self.assertAlmostEqual(
                details[f"total_{clock}_s"],
                details[f"zip_write_{clock}_s"] + details[f"xml_and_other_{clock}_s"], delta=0.002,
            )

    def test_failed_save_does_not_log_a_successful_breakdown(self):
        with patch("excel_save.ExcelWriter.save", side_effect=OSError("disk error")), \
                patch("excel_save.log_excel_event") as event:
            with self.assertRaisesRegex(OSError, "disk error"):
                self._workbook(DiagnosticExcelWriter)
        event.assert_not_called()


if __name__ == "__main__":
    unittest.main()
