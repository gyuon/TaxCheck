import unittest
from io import BytesIO
from unittest.mock import patch

import pandas as pd
from openpyxl import load_workbook

from constants import Col, SheetName, Status
import data_processor


class TestAdjacentAmountExport(unittest.TestCase):
    def test_export_does_not_build_a_series_for_each_payment_row(self):
        errors = pd.DataFrame([{
            Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.FUND_NAME: '운영',
            Col.STATUS: Status.EXCESS, Col.DIFF: 1000,
        }])
        raw = pd.DataFrame([
            {Col.NAME: 'A', Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 1, Col.RAW_DEPOSIT: 500},
            {Col.NAME: 'A', Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 1, Col.RAW_DEPOSIT: 200},
            {Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.CODE1: 1, Col.RAW_DEPOSIT: 999},
            {Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 2, Col.CODE1: 1, Col.RAW_DEPOSIT: 800},
        ])
        original = pd.DataFrame.iterrows

        def guarded(frame):
            if 'canonical' in frame:
                raise AssertionError('전후월 집계에서 행별 Series 생성을 반복함')
            return original(frame)

        with patch.object(pd.DataFrame, 'iterrows', guarded):
            data, _ = data_processor.to_excel_bytes(pd.DataFrame(), errors, 'input_26.11.4.xlsx', df_raw=raw)
        sheet = load_workbook(BytesIO(data))[SheetName.ERROR_RESULT]
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual(sheet.cell(2, headers.index(Col.PREV_MONTH_AMT) + 1).value, 700)
        self.assertEqual(sheet.cell(2, headers.index(Col.NEXT_MONTH_AMT) + 1).value, 800)

    def test_all_statuses_funds_and_year_boundaries_preserve_input_and_order(self):
        rows = []
        raw = []
        expected = []
        for fund, code in [('운영', 1), ('협력', 2), ('복지', 3)]:
            for status in [Status.UNPAID, Status.EXCESS, Status.INSUFFICIENT]:
                rows.append({Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.FUND_NAME: fund, Col.STATUS: status})
                expected.append(code * 100)
            raw.append({Col.NAME: 'A', Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: code, Col.RAW_DEPOSIT: code * 100})
        errors = pd.DataFrame(rows, index=[90, 20, 70, 10, 50, 30, 80, 60, 40])
        source = pd.DataFrame(raw).astype({Col.YEAR: 'Int64', Col.MONTH: 'Int64', Col.CODE1: 'Int64', Col.RAW_DEPOSIT: 'Int64'})
        old_errors, old_source = errors.copy(deep=True), source.copy(deep=True)
        with patch.object(pd.DataFrame, 'iterrows', side_effect=AssertionError('행별 Series 생성')):
            result = data_processor._calc_adjacent_month_amounts(errors, source)
        self.assertEqual(result.index.tolist(), errors.index.tolist())
        self.assertEqual(result[Col.PREV_MONTH_AMT].tolist(), expected)
        self.assertEqual(result[Col.NEXT_MONTH_AMT].tolist(), [0] * 9)
        pd.testing.assert_frame_equal(errors, old_errors)
        pd.testing.assert_frame_equal(source, old_source)

    def test_unknown_fund_or_status_keeps_existing_values(self):
        errors = pd.DataFrame([
            {Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.FUND_NAME: '기타', Col.STATUS: Status.UNPAID},
            {Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.FUND_NAME: '운영', Col.STATUS: '정상'},
        ])
        errors[Col.PREV_MONTH_AMT] = [7, 8]
        errors[Col.NEXT_MONTH_AMT] = [9, 10]
        source = pd.DataFrame([{Col.NAME: 'A', Col.YEAR: 2025, Col.MONTH: 12, Col.CODE1: 1, Col.RAW_DEPOSIT: 100}])
        result = data_processor._calc_adjacent_month_amounts(errors, source)
        self.assertEqual(result[Col.PREV_MONTH_AMT].tolist(), [7, 8])
        self.assertEqual(result[Col.NEXT_MONTH_AMT].tolist(), [9, 10])

    def test_no_source_or_no_errors_keeps_existing_output(self):
        errors = pd.DataFrame([{Col.NAME: 'A'}])
        for source in [None, pd.DataFrame(), pd.DataFrame([{Col.NAME: 'A'}])]:
            with self.subTest(source_type=type(source).__name__):
                self.assertIs(data_processor._calc_adjacent_month_amounts(errors, source), errors)
        empty = pd.DataFrame()
        self.assertIs(data_processor._calc_adjacent_month_amounts(empty, pd.DataFrame([{Col.CODE1: 1}])), empty)


if __name__ == '__main__':
    unittest.main()
