import unittest
from io import BytesIO

import pandas as pd
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

import data_processor
from constants import Col, RESULT_COLUMNS, SheetName, Status


class TestResultFlags(unittest.TestCase):
    def test_error_target_marks_unpaid_and_thousand_won_boundaries(self):
        rows = []
        for name, cooperation_amount in {
            "PositiveBoundary": 41000,
            "PositiveInside": 40999,
            "NegativeBoundary": 39000,
            "NegativeInside": 39001,
        }.items():
            rows.extend([
                {
                    Col.NAME: name,
                    Col.YEAR: 2021,
                    Col.MONTH: 1,
                    Col.CODE1: 1,
                    Col.CODE2: 1,
                    Col.RAW_DEPOSIT: 100000,
                },
                {
                    Col.NAME: name,
                    Col.YEAR: 2021,
                    Col.MONTH: 1,
                    Col.CODE1: 2,
                    Col.CODE2: 1,
                    Col.RAW_DEPOSIT: cooperation_amount,
                },
            ])

        errors = data_processor.detect_errors(pd.DataFrame(rows)).set_index(Col.NAME)

        self.assertEqual(errors.loc["PositiveBoundary", Col.DIFF], 1000)
        self.assertEqual(errors.loc["PositiveBoundary", Col.ERROR_TARGET], "O")
        self.assertEqual(errors.loc["NegativeBoundary", Col.DIFF], -1000)
        self.assertEqual(errors.loc["NegativeBoundary", Col.ERROR_TARGET], "O")
        self.assertEqual(errors.loc["PositiveInside", Col.ERROR_TARGET], "")
        self.assertEqual(errors.loc["NegativeInside", Col.ERROR_TARGET], "")

        unpaid_source = pd.DataFrame([
            {
                Col.NAME: "Unpaid",
                Col.YEAR: 2021,
                Col.MONTH: 1,
                Col.CODE1: 1,
                Col.CODE2: 1,
                Col.RAW_DEPOSIT: 100000,
            }
        ])
        missed, _ = data_processor.generate_missed_months(unpaid_source)
        self.assertTrue((missed[Col.ERROR_TARGET] == "O").all())

    def test_three_month_unpaid_flag_keeps_existing_boundary_rule(self):
        cases = [
            (2026, 1, "※정리본_26.4.14.xlsx", "O"),
            (2026, 2, "※정리본_26.4.14.xlsx", ""),
            (2025, 10, "※정리본_26.1.14.xlsx", "O"),
        ]

        for year, month, filename, expected in cases:
            with self.subTest(year=year, month=month, filename=filename):
                source = pd.DataFrame([
                    {
                        Col.NAME: "Unpaid",
                        Col.YEAR: year,
                        Col.MONTH: month,
                        Col.CODE1: 1,
                        Col.CODE2: 1,
                        Col.RAW_DEPOSIT: 100000,
                    }
                ])
                missed, _ = data_processor.generate_missed_months(source, filename=filename)

                self.assertFalse(missed.empty)
                self.assertTrue((missed[Col.UNPAID_3_MONTHS] == expected).all())

    def test_error_result_sheet_has_new_columns_and_active_target_filter(self):
        df_first = pd.DataFrame({Col.NAME: ["A"], Col.YEAR: [2021], Col.MONTH: [1]})
        source = pd.DataFrame([
            {
                Col.NAME: "A",
                Col.YEAR: 2021,
                Col.MONTH: 1,
                Col.CODE1: 1,
                Col.CODE2: 1,
                Col.RAW_DEPOSIT: 100000,
            },
            {
                Col.NAME: "A",
                Col.YEAR: 2021,
                Col.MONTH: 1,
                Col.CODE1: 2,
                Col.CODE2: 1,
                Col.RAW_DEPOSIT: 41000,
            },
        ])
        errors = data_processor.detect_errors(source)

        excel_bytes, _ = data_processor.to_excel_bytes(
            df_first,
            errors,
            "result.xlsx",
            df_raw=pd.DataFrame(),
        )
        workbook = load_workbook(BytesIO(excel_bytes))
        worksheet = workbook[SheetName.ERROR_RESULT]
        headers = [cell.value for cell in worksheet[1]]

        self.assertEqual(headers, ["번호", *RESULT_COLUMNS])
        self.assertEqual(headers[-1], Col.ERROR_TARGET)
        self.assertEqual(headers[-2], Col.UNPAID_3_MONTHS)
        self.assertIn(Col.ERROR_TARGET, headers)
        self.assertIn(Col.UNPAID_3_MONTHS, headers)
        self.assertNotIn("비고", headers)
        self.assertEqual(
            worksheet.auto_filter.ref,
            f"A1:{get_column_letter(worksheet.max_column)}{worksheet.max_row}",
        )
        filters = worksheet.auto_filter.filterColumn
        self.assertEqual(len(filters), 1)
        self.assertEqual(filters[0].colId, headers.index(Col.ERROR_TARGET))
        self.assertEqual(filters[0].colId, worksheet.max_column - 1)
        self.assertEqual(filters[0].filters.filter, ["O"])
        self.assertTrue(worksheet.sheet_properties.filterMode)
        self.assertFalse(worksheet.row_dimensions[1].hidden)
        self.assertFalse(worksheet.row_dimensions[2].hidden)

        error_target_column = headers.index(Col.ERROR_TARGET) + 1
        long_unpaid_column = headers.index(Col.UNPAID_3_MONTHS) + 1
        self.assertEqual(worksheet.cell(2, error_target_column).value, "O")
        self.assertEqual(worksheet.cell(2, error_target_column).alignment.horizontal, "center")
        self.assertEqual(worksheet.cell(2, long_unpaid_column).alignment.horizontal, "center")
        comment = worksheet.cell(1, error_target_column).comment
        self.assertIsNotNone(comment)
        self.assertEqual(comment.text, (
            '다음 중 하나에 해당하면 오류 대상으로 표시합니다.\n\n'
            '(1) 상태가 "미납"인 모든 항목\n'
            '(2) 차액이 1,000원 이상 또는 -1,000원 이하인 항목\n\n'
            '※ 납부금액의 천 원 미만 절삭 허용 규칙에 따라,\n'
            '   차액이 -999원 ~ 999원인 부족·초과 항목은\n'
            '   오류 대상에서 제외합니다.\n'
            '   단, 미납 항목은 차액과 관계없이 포함합니다.'
        ))
        self.assertEqual(comment.author, "TaxCheck")
        self.assertFalse(any(cell.comment for cell in worksheet[1][:-1]))

        last_column = get_column_letter(worksheet.max_column)
        previous_column = get_column_letter(worksheet.max_column - 1)
        self.assertIn(
            f"{previous_column}2:{last_column}{worksheet.max_row}",
            [str(formatting.sqref) for formatting in worksheet.conditional_formatting],
        )

    def test_initial_filter_hides_only_non_targets_without_removing_data(self):
        cases = {
            "mixed": [
                (Status.UNPAID, None, "O"),
                (Status.EXCESS, 1000, "O"),
                (Status.INSUFFICIENT, -1000, "O"),
                (Status.EXCESS, 999, ""),
                (Status.INSUFFICIENT, -999, ""),
                (Status.EXCESS, 500, pd.NA),
            ],
            "all_targets": [(Status.UNPAID, None, "O"), (Status.EXCESS, 1000, "O")],
            "no_targets": [(Status.EXCESS, 999, ""), (Status.INSUFFICIENT, -999, "")],
        }
        for name, rows in cases.items():
            with self.subTest(name=name):
                errors = pd.DataFrame([
                    {
                        Col.NAME: f"Person{i}", Col.YEAR: 2021, Col.MONTH: 1,
                        Col.FUND_NAME: "협력", Col.CODE: "21", Col.STATUS: status,
                        Col.DIFF: difference, Col.ERROR_TARGET: target,
                    }
                    for i, (status, difference, target) in enumerate(rows)
                ])
                original = errors.copy(deep=True)
                first = pd.DataFrame({Col.NAME: ["Person0"], Col.YEAR: [2021], Col.MONTH: [1]})
                excel_bytes, _ = data_processor.to_excel_bytes(first, errors, "result.xlsx")
                workbook = load_workbook(BytesIO(excel_bytes))
                worksheet = workbook[SheetName.ERROR_RESULT]
                headers = [cell.value for cell in worksheet[1]]
                target_column = headers.index(Col.ERROR_TARGET) + 1
                self.assertEqual(target_column, worksheet.max_column)
                name_column = headers.index(Col.NAME) + 1

                self.assertEqual(worksheet.max_row, len(rows) + 1)
                self.assertEqual(
                    [worksheet.cell(i, name_column).value for i in range(2, worksheet.max_row + 1)],
                    errors[Col.NAME].tolist(),
                )
                self.assertFalse(worksheet.row_dimensions[1].hidden)
                for row_number in range(2, worksheet.max_row + 1):
                    is_target = worksheet.cell(row_number, target_column).value == "O"
                    self.assertEqual(worksheet.row_dimensions[row_number].hidden, not is_target)
                for sheet in workbook:
                    if sheet.title != SheetName.ERROR_RESULT:
                        self.assertFalse(sheet.sheet_properties.filterMode)
                        self.assertEqual(sheet.auto_filter.filterColumn, [])
                        self.assertFalse(any(row.hidden for row in sheet.row_dimensions.values()))
                pd.testing.assert_frame_equal(errors, original)

    def test_excel_export_migrates_legacy_result_columns(self):
        legacy_errors = pd.DataFrame([
            {
                Col.NAME: "Legacy",
                Col.YEAR: 2021,
                Col.MONTH: 1,
                Col.FUND_NAME: "협력",
                Col.CODE: "21",
                Col.STATUS: Status.EXCESS,
                Col.DIFF: 1000,
                Col.DEPOSIT: 41000,
                Col.STANDARD: 40000,
                Col.FORMULA: "",
                Col.PREV_MONTH_AMT: 0,
                Col.NEXT_MONTH_AMT: 0,
                "비고": "3개월 이전 미납분",
            }
        ])

        excel_bytes, _ = data_processor.to_excel_bytes(
            pd.DataFrame({Col.NAME: ["Legacy"], Col.YEAR: [2021], Col.MONTH: [1]}),
            legacy_errors,
            "legacy.xlsx",
            df_raw=pd.DataFrame(),
        )
        worksheet = load_workbook(BytesIO(excel_bytes))[SheetName.ERROR_RESULT]
        headers = [cell.value for cell in worksheet[1]]

        self.assertEqual(headers, ["번호", *RESULT_COLUMNS])
        self.assertEqual(headers[-1], Col.ERROR_TARGET)
        self.assertIsNotNone(worksheet.cell(1, worksheet.max_column).comment)
        self.assertEqual(worksheet.cell(2, headers.index(Col.ERROR_TARGET) + 1).value, "O")
        self.assertEqual(worksheet.cell(2, headers.index(Col.UNPAID_3_MONTHS) + 1).value, "O")
        self.assertEqual(worksheet.auto_filter.filterColumn[0].filters.filter, ["O"])
        self.assertFalse(worksheet.row_dimensions[2].hidden)


if __name__ == "__main__":
    unittest.main()
