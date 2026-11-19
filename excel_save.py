"""기본 openpyxl 저장 방식을 유지하며 ZIP 쓰기와 나머지 저장 시간을 구분한다."""

from datetime import datetime, timezone
import mmap
import time
from zipfile import ZIP_DEFLATED, ZipFile

from openpyxl.writer.excel import ExcelWriter
from pandas.io.excel._openpyxl import OpenpyxlWriter

from excel_timing import log_excel_event


class _TimedZipFile(ZipFile):
    def __init__(self, *args, **kwargs):
        self.wall_s = 0.0
        self.cpu_s = 0.0
        self.write_calls = 0
        super().__init__(*args, **kwargs)

    def _measure(self, callback, *args, **kwargs):
        wall, cpu = time.perf_counter(), time.thread_time()
        try:
            return callback(*args, **kwargs)
        finally:
            self.wall_s += time.perf_counter() - wall
            self.cpu_s += time.thread_time() - cpu

    def write(self, *args, **kwargs):
        self.write_calls += 1
        return self._measure(super().write, *args, **kwargs)

    def writestr(self, *args, **kwargs):
        self.write_calls += 1
        return self._measure(super().writestr, *args, **kwargs)

    def close(self):
        if getattr(self, "fp", None) is not None:
            return self._measure(super().close)


class DiagnosticExcelWriter(OpenpyxlWriter):
    def _save(self):
        # pandas/OpenpyxlWriter._save와 openpyxl/save_workbook의 기본 절차를 유지한다.
        wall, cpu = time.perf_counter(), time.thread_time()
        with _TimedZipFile(self._handles.handle, "w", ZIP_DEFLATED, allowZip64=True) as archive:
            self.book.properties.modified = datetime.now(timezone.utc).replace(tzinfo=None)
            ExcelWriter(self.book, archive).save()
        if "r+" in self._mode and not isinstance(self._handles.handle, mmap.mmap):
            self._handles.handle.truncate()
        total_wall, total_cpu = time.perf_counter() - wall, time.thread_time() - cpu
        log_excel_event(
            "save_breakdown", total_wall_s=round(total_wall, 3), total_cpu_s=round(total_cpu, 3),
            zip_write_wall_s=round(archive.wall_s, 3), zip_write_cpu_s=round(archive.cpu_s, 3),
            xml_and_other_wall_s=round(max(0, total_wall - archive.wall_s), 3),
            xml_and_other_cpu_s=round(max(0, total_cpu - archive.cpu_s), 3),
            zip_write_calls=archive.write_calls,
            uncompressed_bytes=sum(entry.file_size for entry in archive.filelist),
            compressed_bytes=sum(entry.compress_size for entry in archive.filelist),
        )
