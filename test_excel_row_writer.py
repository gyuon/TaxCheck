import unittest
from datetime import date, datetime, timedelta, timezone
from io import BytesIO
from unittest.mock import patch

import numpy as np
import pandas as pd
from openpyxl import load_workbook

import data_processor
from constants import Col, SheetName, Status


class TestResultRowWriter(unittest.TestCase):
    @staticmethod
    def write_frame(frame, direct):
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
            if direct:
                data_processor._write_error_result_rows(writer, frame)
            else:
                frame.to_excel(writer, index=True, index_label='번호', sheet_name=SheetName.ERROR_RESULT)
        return load_workbook(BytesIO(buffer.getvalue()))[SheetName.ERROR_RESULT]

    def test_scalar_values_types_and_number_formats_match_pandas(self):
        values = [
            None, pd.NA, np.nan, pd.NaT, np.inf, -np.inf,
            np.int64(12), np.float32(1.25), np.bool_(True), False,
            '=SUM(1,2)', '문자열', '', 'a\nb',
            date(2026, 1, 2), datetime(2026, 1, 2, 3, 4), timedelta(days=1, hours=12),
            pd.Timestamp('2026-02-03'), pd.Timedelta(minutes=30), {'k': 1}, [1, 2],
            np.datetime64('2026-01-01'), np.timedelta64(1, 'D'),
        ]
        frame = pd.DataFrame({'값': pd.Series(values, dtype=object)})
        original = frame.copy(deep=True)
        before, after = self.write_frame(frame, False), self.write_frame(frame, True)
        self.assertEqual(before.max_row, after.max_row)
        for old_row, new_row in zip(before, after):
            for old, new in zip(old_row, new_row):
                with self.subTest(cell=old.coordinate):
                    self.assertEqual((old.value, old.data_type, old.number_format),
                                     (new.value, new.data_type, new.number_format))
        pd.testing.assert_frame_equal(frame, original)

    def test_nullable_columns_keep_blanks_and_values(self):
        frame = pd.DataFrame({
            '정수': pd.array([1, None, -2], dtype='Int64'),
            '실수': pd.array([1.25, None, -2.5], dtype='Float64'),
            '논리': pd.array([True, None, False], dtype='boolean'),
            '문자': pd.array(['A', None, '=1+2'], dtype='string'),
        })
        before, after = self.write_frame(frame, False), self.write_frame(frame, True)
        for old_row, new_row in zip(before, after):
            self.assertEqual([(c.value, c.data_type) for c in old_row], [(c.value, c.data_type) for c in new_row])

    def test_timezone_dates_are_rejected_as_before(self):
        frame = pd.DataFrame({'날짜': [datetime(2026, 1, 2, tzinfo=timezone.utc)]})
        for direct in [False, True]:
            with self.subTest(direct=direct), self.assertRaisesRegex(ValueError, 'timezones'):
                self.write_frame(frame, direct)

    def test_result_export_uses_row_writing_and_keeps_filter_and_comment(self):
        errors = pd.DataFrame([
            {Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.FUND_NAME: '협력', Col.STATUS: Status.EXCESS,
             Col.DIFF: 1000, Col.ERROR_TARGET: 'O', Col.FORMULA: '=1+2'},
            {Col.NAME: 'B', Col.YEAR: 2026, Col.MONTH: 2, Col.FUND_NAME: '복지', Col.STATUS: Status.INSUFFICIENT,
             Col.DIFF: -999, Col.ERROR_TARGET: '', Col.FORMULA: None},
        ])
        original = pd.DataFrame.to_excel

        def guarded(frame, *args, **kwargs):
            self.assertNotEqual(kwargs.get('sheet_name'), SheetName.ERROR_RESULT)
            return original(frame, *args, **kwargs)

        with patch.object(pd.DataFrame, 'to_excel', guarded):
            data, _ = data_processor.to_excel_bytes(pd.DataFrame(), errors, 'input_26.11.4.xlsx')
        sheet = load_workbook(BytesIO(data))[SheetName.ERROR_RESULT]
        headers = [cell.value for cell in sheet[1]]
        self.assertEqual(sheet['A2'].value, 1)
        self.assertEqual(sheet['A3'].value, 2)
        formula = sheet.cell(2, headers.index(Col.FORMULA) + 1)
        self.assertEqual((formula.value, formula.data_type), ('=1+2', 'f'))
        self.assertFalse(sheet.row_dimensions[2].hidden)
        self.assertTrue(sheet.row_dimensions[3].hidden)
        self.assertEqual(sheet.auto_filter.filterColumn[0].filters.filter, ['O'])
        self.assertIsNotNone(sheet.cell(1, len(headers)).comment)


if __name__ == '__main__':
    unittest.main()
