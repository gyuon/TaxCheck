import unittest
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.units import points_to_pixels

from constants import Col, SheetName, Status
from data_processor import to_excel_bytes
from excess_shortfall import filter_comparison_period
from unpaid import UNPAID_FUNDS, build_unpaid_summary


def record(name="홍길동", year=2026, month=8, fund="운영", status=Status.UNPAID, target="O"):
    return {
        Col.NAME: name, Col.YEAR: year, Col.MONTH: month, Col.FUND_NAME: fund,
        Col.CODE: "11~17", Col.STATUS: status, Col.ERROR_TARGET: target,
        Col.STANDARD: 100000, Col.DIFF: -100000, Col.DEPOSIT: 0,
    }


class TestComparisonPeriod(unittest.TestCase):
    def test_cutoff_includes_three_months_before_and_preserves_input(self):
        errors = pd.DataFrame([record(month=month) for month in [7, 8, 9, 10, 11]])
        original = errors.copy(deep=True)
        filtered = filter_comparison_period(errors, (2026, 11))
        self.assertEqual(filtered[Col.MONTH].tolist(), [7, 8])
        pd.testing.assert_frame_equal(errors, original)

    def test_year_boundary_and_analysis_range(self):
        errors = pd.DataFrame([
            record(year=2024, month=12), record(year=2025, month=10),
            record(year=2025, month=11), record(year=2026, month=1),
        ])
        filtered = filter_comparison_period(errors, (2026, 1), 2025, 2026)
        self.assertEqual(list(zip(filtered[Col.YEAR], filtered[Col.MONTH])), [(2025, 10)])

    def test_missing_or_invalid_reference_month_is_rejected(self):
        for reference in [None, (2026, 0), (2026, 13)]:
            with self.subTest(reference=reference):
                with self.assertRaisesRegex(ValueError, "기준 년월"):
                    filter_comparison_period(pd.DataFrame([record()]), reference)

    def test_both_sheets_share_cutoff_without_changing_other_results(self):
        errors = pd.DataFrame([
            record(month=8), record(month=9),
            record(name="초과대상", month=8, status=Status.EXCESS),
            record(name="초과대상", month=9, status=Status.EXCESS),
        ])
        data, _ = to_excel_bytes(pd.DataFrame(), errors, "정리본_26.11.4_오류검출.xlsx")
        workbook = load_workbook(BytesIO(data))
        self.assertEqual(workbook.sheetnames[:3], [SheetName.ERROR_RESULT, SheetName.EXCESS_SHORTFALL, SheetName.UNPAID])
        self.assertEqual(workbook[SheetName.ERROR_RESULT].max_row, 5)
        self.assertEqual(workbook[SheetName.EXCESS_SHORTFALL]["B1"].value, "2026-08")
        self.assertEqual(workbook[SheetName.EXCESS_SHORTFALL].max_row, 3)
        self.assertEqual(workbook[SheetName.UNPAID]["B2"].value, "2026-08")
        self.assertEqual(workbook[SheetName.UNPAID].max_row, 2)
        self.assertIsNotNone(workbook[SheetName.ERROR_RESULT]["O1"].comment)

    def test_explicit_source_filename_is_used_and_invalid_name_is_not_guessed(self):
        data, _ = to_excel_bytes(pd.DataFrame(), pd.DataFrame([record()]), "result.xlsx",
                                 source_filename="정리본_26.11.4.xlsx")
        self.assertEqual(load_workbook(BytesIO(data))[SheetName.UNPAID]["B2"].value, "2026-08")
        with self.assertRaisesRegex(ValueError, "기준 년월"):
            to_excel_bytes(pd.DataFrame(), pd.DataFrame([record()]), "result.xlsx", source_filename="unknown.xlsx")


class TestUnpaidSheet(unittest.TestCase):
    def test_groups_months_deduplicates_funds_and_sorts_without_mutation(self):
        errors = pd.DataFrame([
            record(name="홍순이", fund="복지"), record(name="홍순이", fund="협력"),
            record(name="홍순이", fund="복지"), record(name="홍순이", fund="운영"),
            record(month=7), record(month=8),
            record(name="제외", status=Status.EXCESS), record(name="제외", target=""),
        ])
        original = errors.copy(deep=True)
        summary = build_unpaid_summary(errors)
        self.assertEqual(summary[Col.NAME].tolist(), ["홍길동", "홍길동", "홍순이"])
        self.assertEqual(summary[Col.MONTH].tolist(), [8, 7, 8])
        self.assertEqual(summary[UNPAID_FUNDS].tolist(), ["운영", "운영", "운영, 협력, 복지"])
        pd.testing.assert_frame_equal(errors, original)

    def test_layout_colors_merging_sizes_and_literal_names(self):
        errors = pd.DataFrame([
            record(month=7), record(month=8), record(name="홍순이", fund="복지"),
            record(name="홍순이", fund="협력"),
        ])
        data, _ = to_excel_bytes(pd.DataFrame(), errors, "result.xlsx", source_filename="정리본_26.11.4.xlsx")
        sheet = load_workbook(BytesIO(data))[SheetName.UNPAID]
        self.assertEqual([cell.value for cell in sheet[1]], ["이름", "년월", "미납 기금명"])
        self.assertEqual(sheet["B2"].value, "2026-08")
        self.assertEqual(sheet["B3"].value, "2026-07")
        self.assertEqual(sheet["A5"].value, "홍순이")
        self.assertEqual(sheet["C5"].value, "협력, 복지")
        self.assertIn("A2:A3", [str(item) for item in sheet.merged_cells.ranges])
        self.assertEqual([sheet.cell(2, col).fill.fgColor.rgb for col in [1, 2, 3]],
                         ["FFFFFFA7", "FFD4E3F8", "FFF8E3D4"])
        self.assertEqual([sheet.column_dimensions[col].width for col in "ABC"], [11, 11, 22])
        for row_number in range(1, 6):
            self.assertEqual(points_to_pixels(sheet.row_dimensions[row_number].height, dpi=120),
                             28 if row_number == 4 else 36)
        for row in sheet:
            for cell in row:
                if cell.value is not None:
                    self.assertEqual(cell.alignment.horizontal, "center")
                    self.assertEqual(cell.alignment.vertical, "center")
        self.assertIsNone(sheet.auto_filter.ref)
        self.assertEqual(sheet.freeze_panes, "A2")
        name = '=HYPERLINK("https://example.invalid")'
        data, _ = to_excel_bytes(pd.DataFrame(), pd.DataFrame([record(name=name)]), "result.xlsx",
                                 source_filename="정리본_26.11.4.xlsx")
        cell = load_workbook(BytesIO(data))[SheetName.UNPAID]["A2"]
        self.assertEqual(cell.value, name)
        self.assertEqual(cell.data_type, "s")

    def test_empty_sheet_and_no_qualifying_month_have_message(self):
        for errors in [pd.DataFrame(), pd.DataFrame([record(month=9)])]:
            data, _ = to_excel_bytes(pd.DataFrame(), errors, "result.xlsx", source_filename="정리본_26.11.4.xlsx")
            sheet = load_workbook(BytesIO(data))[SheetName.UNPAID]
            self.assertEqual(sheet["A2"].value, "미납 항목이 없습니다.")


if __name__ == "__main__":
    unittest.main()
