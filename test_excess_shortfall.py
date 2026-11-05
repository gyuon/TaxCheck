import unittest
import math
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils.units import points_to_pixels

from constants import Col, FundName, SheetName, Status
from data_processor import to_excel_bytes
from excess_shortfall import build_excess_shortfall_summary, write_excess_shortfall_sheet


def error(name="김선애", year=2026, month=11, fund=FundName.COOPERATION,
          status=Status.INSUFFICIENT, target="O", deposit=523000, standard=1200000):
    return {
        Col.NAME: name, Col.YEAR: year, Col.MONTH: month, Col.FUND_NAME: fund,
        Col.CODE: "21" if fund == FundName.COOPERATION else "31",
        Col.STATUS: status, Col.ERROR_TARGET: target, Col.DEPOSIT: deposit,
        Col.STANDARD: standard, Col.DIFF: deposit - standard,
    }


def payment(name="김선애", year=2026, month=11, code=1, amount=100000):
    return {
        Col.NAME: name, Col.YEAR: year, Col.MONTH: month,
        Col.CODE1: code, Col.CODE2: 1, Col.RAW_DEPOSIT: amount,
    }


class TestExcessShortfallSummary(unittest.TestCase):
    def test_selects_only_target_excess_and_insufficient_without_month_cutoff(self):
        errors = pd.DataFrame([
            error(month=11),
            error(month=10, status=Status.EXCESS),
            error(month=8),
            error(name="미납제외", status=Status.UNPAID),
            error(name="소액제외", target="", deposit=1200999),
            error(name="빈값제외", target=pd.NA),
        ])

        summary = build_excess_shortfall_summary(errors)

        self.assertEqual(summary[Col.NAME].tolist(), ["김선애"] * 3)
        self.assertEqual(summary[Col.MONTH].tolist(), [11, 10, 8])

    def test_sums_all_funds_including_non_error_payments_without_mutation(self):
        errors = pd.DataFrame([error()])
        raw = pd.DataFrame([
            payment(amount=1500000), payment(amount=1604000),
            payment(code=2, amount=523000), payment(code=3, amount=3104000),
            payment(name="다른사람", amount=999999),
            payment(month=10, amount=999999), payment(code=9, amount=999999),
        ])
        original_errors, original_raw = errors.copy(deep=True), raw.copy(deep=True)

        row = build_excess_shortfall_summary(errors, raw).iloc[0]

        self.assertEqual(row[f"운영_{Col.DEPOSIT}"], 3104000)
        self.assertEqual(row[f"협력_{Col.DEPOSIT}"], 523000)
        self.assertEqual(row[f"복지_{Col.DEPOSIT}"], 3104000)
        self.assertEqual(row[f"협력_{Col.STANDARD}"], 1200000)
        self.assertTrue(pd.isna(row[f"운영_{Col.STANDARD}"]))
        self.assertTrue(pd.isna(row[f"복지_{Col.STANDARD}"]))
        pd.testing.assert_frame_equal(errors, original_errors)
        pd.testing.assert_frame_equal(raw, original_raw)

    def test_multiple_error_funds_share_one_month_but_non_target_standard_is_blank(self):
        errors = pd.DataFrame([
            error(),
            error(fund=FundName.WELFARE, status=Status.EXCESS, standard=3104000),
            error(month=10),
            error(month=10, fund=FundName.WELFARE, target="", deposit=3104999, standard=3104000),
        ])

        summary = build_excess_shortfall_summary(errors)

        self.assertEqual(len(summary), 2)
        self.assertEqual(summary.iloc[0][f"협력_{Col.STANDARD}"], 1200000)
        self.assertEqual(summary.iloc[0][f"복지_{Col.STANDARD}"], 3104000)
        self.assertTrue(pd.isna(summary.iloc[1][f"복지_{Col.STANDARD}"]))

    def test_absent_payment_is_zero_when_raw_records_are_available(self):
        raw = pd.DataFrame([payment(code=2, amount=523000)])

        row = build_excess_shortfall_summary(pd.DataFrame([error()]), raw).iloc[0]

        self.assertEqual(row[f"운영_{Col.DEPOSIT}"], 0)
        self.assertEqual(row[f"복지_{Col.DEPOSIT}"], 0)

    def test_without_raw_unknown_payments_remain_blank_not_zero(self):
        row = build_excess_shortfall_summary(pd.DataFrame([error()])).iloc[0]

        self.assertEqual(row[f"협력_{Col.DEPOSIT}"], 523000)
        self.assertTrue(pd.isna(row[f"운영_{Col.DEPOSIT}"]))
        self.assertTrue(pd.isna(row[f"복지_{Col.DEPOSIT}"]))

    def test_empty_or_excluded_data_has_no_comparison_rows(self):
        for errors in [None, pd.DataFrame(), pd.DataFrame([error(status=Status.UNPAID), error(target="")])]:
            with self.subTest(errors=errors):
                self.assertTrue(build_excess_shortfall_summary(errors).empty)


class TestExcessShortfallSheet(unittest.TestCase):
    def test_sizes_match_reported_excel_environment_at_120_dpi(self):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            write_excess_shortfall_sheet(writer, pd.DataFrame([error()]))
        sheet = load_workbook(BytesIO(buffer.getvalue()))[SheetName.EXCESS_SHORTFALL]

        # 실제 Excel에서 관측한 최대 숫자 폭 10px에 OOXML의 너비 환산식을 적용한다.
        for column, pixels in {"A": 110, "B": 110, "C": 120, "D": 120, "E": 120}.items():
            width = sheet.column_dimensions[column].width
            displayed_pixels = math.trunc(((256 * width + math.trunc(128 / 10)) / 256) * 10)
            self.assertEqual(displayed_pixels, pixels)
        self.assertAlmostEqual(sheet.row_dimensions[1].height, 21.6)
        self.assertEqual(points_to_pixels(sheet.row_dimensions[1].height, dpi=120), 36)
        self.assertEqual(sheet.sheet_view.zoomScale, 100)

    def test_layout_sorting_merges_amounts_and_styles(self):
        errors = pd.DataFrame([
            error(name="서윤구"),
            error(month=8, fund=FundName.WELFARE, status=Status.EXCESS, deposit=3999000, standard=3140000),
            error(),
            error(fund=FundName.WELFARE, status=Status.EXCESS, deposit=3999000, standard=3104000),
        ])
        raw = pd.DataFrame([
            payment(amount=1500000), payment(amount=1604000),
            payment(code=2, amount=523000), payment(code=3, amount=3999000),
            payment(month=8, amount=3104000), payment(month=8, code=2, amount=523000),
            payment(month=8, code=3, amount=3999000),
            payment(name="서윤구"), payment(name="서윤구", code=2, amount=523000),
        ])
        data, filename = to_excel_bytes(pd.DataFrame(), errors, "※정리본_26.11.4.xlsx", df_raw=raw)
        workbook = load_workbook(BytesIO(data))
        sheet = workbook[SheetName.EXCESS_SHORTFALL]

        self.assertEqual(filename, "※정리본_26.11.4.xlsx")
        self.assertEqual(workbook.sheetnames[:2], [SheetName.ERROR_RESULT, SheetName.EXCESS_SHORTFALL])
        self.assertEqual(sheet.max_column, 5)
        self.assertEqual(sheet.max_row, 10)
        self.assertEqual(sheet["A1"].value, "김선애")
        self.assertEqual(sheet["A8"].value, "서윤구")
        self.assertEqual(sheet["B1"].value, "2026-11")
        self.assertEqual(sheet["B4"].value, "2026-08")
        self.assertEqual(set(str(item) for item in sheet.merged_cells.ranges), {
            "A1:A6", "A8:A10", "B1:B3", "B4:B6", "B8:B10",
        })
        self.assertEqual([sheet.cell(1, column).value for column in range(3, 6)], ["운영", "협력", "복지"])
        self.assertEqual([sheet.cell(2, column).value for column in range(3, 6)], [3104000, 523000, 3999000])
        self.assertEqual([sheet.cell(3, column).value for column in range(3, 6)], [None, 1200000, 3104000])
        self.assertEqual([sheet.cell(6, column).value for column in range(3, 6)], [None, None, 3140000])
        self.assertTrue(all(sheet.cell(7, column).value is None for column in range(1, 6)))
        self.assertEqual(sheet["A1"].fill.fgColor.rgb, "FFFFFFA7")
        self.assertEqual(sheet["A8"].fill.fgColor.rgb, "FFFFFFA7")
        for row_number in [1, 4, 8]:
            self.assertEqual(sheet.cell(row_number, 2).fill.fgColor.rgb, "FFD4E3F8")
            for column in range(3, 6):
                self.assertEqual(sheet.cell(row_number, column).fill.fgColor.rgb, "FFF8E3D4")
                self.assertIsNone(sheet.cell(row_number + 1, column).fill.fill_type)
                self.assertIsNone(sheet.cell(row_number + 2, column).fill.fill_type)
        self.assertEqual(sheet["D3"].font.color.rgb, "FFFF0000")
        self.assertEqual(sheet["D3"].number_format, "#,##0")
        self.assertEqual(sheet["D3"].data_type, "n")
        self.assertEqual(sheet["D3"].alignment.horizontal, "center")
        self.assertEqual(sheet["A1"].alignment.vertical, "center")
        self.assertEqual(sheet["C2"].border.bottom.style, "thin")
        self.assertIsNone(sheet.auto_filter.ref)
        self.assertFalse(any(row.hidden for row in sheet.row_dimensions.values()))
        for column, pixels in {"A": 110, "B": 110, "C": 120, "D": 120, "E": 120}.items():
            width = sheet.column_dimensions[column].width
            self.assertAlmostEqual(width, pixels / 10)
            self.assertEqual(math.trunc(((256 * width + math.trunc(128 / 10)) / 256) * 10), pixels)
        self.assertEqual(points_to_pixels(sheet.sheet_format.defaultRowHeight, dpi=120), 36)
        for row_number in range(1, sheet.max_row + 1):
            expected_pixels = 28 if row_number == 7 else 36
            self.assertEqual(points_to_pixels(sheet.row_dimensions[row_number].height, dpi=120), expected_pixels)
        self.assertAlmostEqual(sheet.row_dimensions[7].height, 16.8)
        for row in sheet.iter_rows():
            for cell in row:
                if cell.value is not None:
                    self.assertEqual(cell.alignment.horizontal, "center")
                    self.assertEqual(cell.alignment.vertical, "center")
        result = workbook[SheetName.ERROR_RESULT]
        self.assertEqual(result.auto_filter.filterColumn[0].filters.filter, ["O"])
        self.assertIsNotNone(result.cell(1, result.max_column).comment)

    def test_name_and_month_order_across_year_boundary(self):
        errors = pd.DataFrame([
            error(name="홍길동", year=2025, month=12),
            error(name="가나다", year=2025, month=12),
            error(name="가나다", year=2026, month=1),
        ])

        summary = build_excess_shortfall_summary(errors)

        self.assertEqual(summary[Col.NAME].tolist(), ["가나다", "가나다", "홍길동"])
        self.assertEqual(list(zip(summary[Col.YEAR], summary[Col.MONTH])), [(2026, 1), (2025, 12), (2025, 12)])
        data, _ = to_excel_bytes(pd.DataFrame(), errors, "result.xlsx")
        sheet = load_workbook(BytesIO(data))[SheetName.EXCESS_SHORTFALL]
        self.assertEqual(sheet["B1"].value, "2026-01")
        self.assertEqual(sheet["B4"].value, "2025-12")

    def test_sheet_exists_with_message_when_no_target_errors(self):
        for errors in [pd.DataFrame(), pd.DataFrame([error(status=Status.UNPAID), error(target="")])]:
            with self.subTest(errors=errors):
                data, _ = to_excel_bytes(pd.DataFrame(), errors, "empty.xlsx")
                sheet = load_workbook(BytesIO(data))[SheetName.EXCESS_SHORTFALL]
                self.assertEqual(sheet["A1"].value, "초과·부족 오류 대상이 없습니다.")
                self.assertEqual(sheet.auto_filter.filterColumn, [])
                self.assertEqual(sheet["A1"].alignment.horizontal, "center")
                self.assertEqual(sheet["A1"].alignment.vertical, "center")
                self.assertEqual(points_to_pixels(sheet.row_dimensions[1].height, dpi=120), 36)

    def test_names_are_written_as_literal_text_not_formulas(self):
        name = '=HYPERLINK("https://example.invalid")'
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            write_excess_shortfall_sheet(writer, pd.DataFrame([error(name=name)]))
        sheet = load_workbook(BytesIO(buffer.getvalue()))[SheetName.EXCESS_SHORTFALL]

        self.assertEqual(sheet["A1"].value, name)
        self.assertEqual(sheet["A1"].data_type, "s")


if __name__ == "__main__":
    unittest.main()
