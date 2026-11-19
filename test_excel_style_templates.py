import unittest
from io import BytesIO
from unittest.mock import patch

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.styles.styleable import StyleDescriptor

from constants import Col, SheetName, Status
from data_processor import to_excel_bytes
from excel_styles import StyleTemplates


class TestStyleTemplates(unittest.TestCase):
    def test_preserves_existing_styles_and_cell_changes_are_independent(self):
        sheet = Workbook().active
        cache = StyleTemplates()
        sheet['A1'].fill = PatternFill('solid', fgColor='FFFF0000')
        for address in ['A1', 'B1', 'C1']:
            cache.apply(sheet[address], 'font', font=Font(bold=True))
        self.assertEqual(sheet['A1'].fill.fgColor.rgb, 'FFFF0000')
        self.assertEqual(sheet['B1'].fill.patternType, None)
        sheet['B1'].number_format = '#,##0'
        self.assertEqual(sheet['C1'].number_format, 'General')
        self.assertTrue(sheet['C1'].font.bold)

    def test_cannot_reuse_style_ids_from_another_workbook(self):
        cache = StyleTemplates()
        cache.apply(Workbook().active['A1'], 'font', font=Font(bold=True))
        with self.assertRaises(ValueError):
            cache.apply(Workbook().active['A1'], 'font', font=Font(bold=True))


class TestExportStyleReuse(unittest.TestCase):
    def test_repeated_result_rows_do_not_reregister_each_style_component(self):
        errors = pd.DataFrame([{
            Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.FUND_NAME: '협력',
            Col.STATUS: Status.EXCESS, Col.DIFF: 1000, Col.DEPOSIT: 41000,
            Col.STANDARD: 40000, Col.CODE: '21', Col.ERROR_TARGET: 'O',
        }] * 200)
        original = StyleDescriptor.__set__
        calls = []

        def count(descriptor, cell, value):
            if cell.parent.title == SheetName.ERROR_RESULT and cell.row and cell.row > 1:
                calls.append(descriptor.key)
            return original(descriptor, cell, value)

        with patch.object(StyleDescriptor, '__set__', count):
            data, _ = to_excel_bytes(pd.DataFrame(), errors, 'input_26.11.4.xlsx', start_year=2026, end_year=2026)
        self.assertLess(len(calls), 1000)
        sheet = load_workbook(BytesIO(data))[SheetName.ERROR_RESULT]
        columns = {cell.value: cell.column for cell in sheet[1]}
        for row in [2, 101, 201]:
            self.assertEqual(sheet.cell(row, columns[Col.STATUS]).font.color.rgb, 'FF059669')
            self.assertEqual(sheet.cell(row, columns[Col.DIFF]).number_format, '+#,##0;-#,##0;0')
            self.assertEqual(sheet.cell(row, columns[Col.DIFF]).alignment.horizontal, 'right')


if __name__ == '__main__':
    unittest.main()
