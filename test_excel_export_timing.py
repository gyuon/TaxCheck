import unittest
from collections import Counter
from io import BytesIO
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.worksheet import Worksheet

from constants import Col, SheetName, Status
from data_processor import to_excel_bytes
import excel_timing
from wide_summary import _style_sheet


class TestWideSummaryPerformance(unittest.TestCase):
    def test_last_row_is_calculated_once_and_bottom_border_is_preserved(self):
        sheet = Workbook().active
        sheet.cell(1, 1, Col.NAME)
        sheet.cell(3, 3, 1000)
        layout = [("field", Col.NAME, Col.NAME), ("sep", None, None), ("total", "총계", "총금액")]
        original = Worksheet.max_row.fget
        calls = []

        def counted_max_row(ws):
            calls.append(ws)
            return original(ws)

        with patch.object(Worksheet, "max_row", property(counted_max_row)):
            _style_sheet(sheet, layout)

        self.assertEqual(len(calls), 1)
        self.assertEqual(sheet["C3"].border.bottom.style, "thick")
        self.assertEqual(sheet["C2"].border.bottom.style, "thin")
        self.assertEqual(sheet.auto_filter.ref, "A2:C3")


class TestExcelStageLogs(unittest.TestCase):
    def test_export_logs_each_stage_and_totals_without_personal_values(self):
        errors = pd.DataFrame([{
            Col.NAME: "Private Name", Col.YEAR: 2026, Col.MONTH: 8,
            Col.FUND_NAME: "협력", Col.CODE: "21", Col.STATUS: Status.EXCESS,
            Col.ERROR_TARGET: "O", Col.DIFF: 1000, Col.DEPOSIT: 41000, Col.STANDARD: 40000,
        }])
        with self.assertLogs("taxcheck", level="INFO") as logs:
            data, _ = to_excel_bytes(
                pd.DataFrame(), errors, "result.xlsx", start_year=2026, end_year=2026,
                source_filename="PrivateFile_26.11.4.xlsx",
            )
        started, completed = [], []
        for record in logs.records:
            message = record.getMessage()
            self.assertNotIn("Private Name", message)
            self.assertNotIn("PrivateFile", message)
            if "event=start" in message:
                started.append(record.args[0])
            if "event=complete" in message:
                completed.append(record.args[0])
                self.assertGreaterEqual(record.args[1], 0)
                self.assertGreaterEqual(record.args[2], 0)
                self.assertIn("peak_rss_mb=", message)
        self.assertEqual(Counter(started), Counter(completed))
        self.assertEqual(started[0], "total")
        self.assertEqual(completed[-1], "total")
        self.assertTrue({
            "adjacent_month_amounts", "error_result_sheet", "excess_shortfall_sheet",
            "unpaid_sheet", "person_summary_aggregation", "fund_summary_aggregation",
            "wide_sheet_data", "wide_sheet_style", "error_summary_sheet",
            "first_payment_sheet", "workbook_save_zip", "comment_position_zip",
        }.issubset(started))
        self.assertEqual(started.count("wide_sheet_style"), 2)
        styled_sheets = {
            record.args[-1]["sheet"] for record in logs.records
            if "event=complete stage=wide_sheet_style " in record.getMessage()
        }
        self.assertEqual(styled_sheets, {SheetName.PERSON_ERROR_STATUS, SheetName.FUND_ERROR_STATUS})
        self.assertIn(SheetName.UNPAID, load_workbook(BytesIO(data)).sheetnames)

    def test_logs_work_when_peak_memory_measurement_is_unavailable(self):
        with patch.object(excel_timing, "resource", None):
            with self.assertLogs("taxcheck", level="INFO") as logs:
                finish = excel_timing.start_excel_stage("portable_stage")
                finish(rows=0)
        self.assertIn("peak_rss_mb=unavailable", logs.output[-1])

    def test_failed_stage_does_not_log_successful_completion(self):
        errors = pd.DataFrame([{
            Col.NAME: "A", Col.YEAR: 2026, Col.MONTH: 8, Col.FUND_NAME: "운영",
            Col.STATUS: Status.UNPAID, Col.ERROR_TARGET: "O", Col.STANDARD: 1000,
        }])
        with self.assertLogs("taxcheck", level="INFO") as logs:
            with self.assertRaisesRegex(ValueError, "기준 년월"):
                to_excel_bytes(pd.DataFrame(), errors, "unknown.xlsx")
        self.assertTrue(any("event=start stage=comparison_period_filter " in message for message in logs.output))
        self.assertFalse(any("event=complete stage=comparison_period_filter " in message for message in logs.output))
        self.assertFalse(any("event=complete stage=total " in message for message in logs.output))


if __name__ == "__main__":
    unittest.main()
