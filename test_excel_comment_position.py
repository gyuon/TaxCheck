import unittest
from io import BytesIO
from xml.etree import ElementTree as ET
from zipfile import ZipFile

import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.comments import Comment

from constants import Col, SheetName, Status
from data_processor import _align_header_comment_right, to_excel_bytes


VML = "urn:schemas-microsoft-com:vml"
EXCEL = "urn:schemas-microsoft-com:office:excel"


class TestExcelCommentPosition(unittest.TestCase):
    def test_exported_note_right_edge_matches_final_column(self):
        errors = pd.DataFrame([{
            Col.NAME: "A", Col.YEAR: 2021, Col.MONTH: 1,
            Col.FUND_NAME: "협력", Col.CODE: "21", Col.STATUS: Status.EXCESS,
            Col.DIFF: 1000, Col.ERROR_TARGET: "O",
        }])
        data, _ = to_excel_bytes(pd.DataFrame(), errors, "result.xlsx")
        worksheet = load_workbook(BytesIO(data))[SheetName.ERROR_RESULT]
        with ZipFile(BytesIO(data)) as archive:
            path = next(name for name in archive.namelist() if name.endswith(".vml"))
            shape = ET.fromstring(archive.read(path)).find(f"{{{VML}}}shape")
        client = shape.find(f"{{{EXCEL}}}ClientData")
        anchor = [int(value) for value in client.findtext(f"{{{EXCEL}}}Anchor").split(",")]
        style = dict(part.split(":", 1) for part in shape.get("style").split(";"))

        self.assertEqual(worksheet.cell(1, worksheet.max_column).value, Col.ERROR_TARGET)
        self.assertEqual(anchor[:4], [10, 146, 0, 2])
        self.assertEqual(anchor[4:6], [worksheet.max_column, 0])
        self.assertEqual(style["margin-left"], "903px")
        self.assertEqual(style["width"], "460px")
        self.assertEqual(style["height"], "200px")
        self.assertIsNone(client.find(f"{{{EXCEL}}}SizeWithCells"))
        self.assertEqual(client.findtext(f"{{{EXCEL}}}Column"), "14")
        self.assertIn("-999원 ~ 999원", worksheet.cell(1, worksheet.max_column).comment.text)
        self.assertEqual(worksheet.auto_filter.filterColumn[0].filters.filter, ["O"])

    def test_other_zip_parts_and_notes_are_unchanged(self):
        workbook = Workbook()
        worksheet = workbook.active
        for name, width in {"A": 8, "B": 10, "C": 12, "D": 14}.items():
            worksheet.column_dimensions[name].width = width
        worksheet.row_dimensions[1].height = 30
        worksheet.row_dimensions[2].hidden = True
        worksheet.row_dimensions[3].hidden = True
        worksheet["D1"].comment = Comment("Header note", "TaxCheck", width=200, height=80)
        worksheet["A2"].comment = Comment("Other note", "Other")
        workbook.create_sheet("Other")["A1"].comment = Comment("Other sheet note", "Other")
        buffer = BytesIO()
        workbook.save(buffer)
        original = buffer.getvalue()
        result = _align_header_comment_right(original, worksheet["D1"])

        with ZipFile(BytesIO(original)) as before, ZipFile(BytesIO(result)) as after:
            self.assertEqual(before.namelist(), after.namelist())
            self.assertIsNone(after.testzip())
            path = worksheet.legacy_drawing
            for name in before.namelist():
                if name != path:
                    self.assertEqual(before.read(name), after.read(name), name)
            original_shapes = ET.fromstring(before.read(path)).findall(f"{{{VML}}}shape")
            updated_shapes = ET.fromstring(after.read(path)).findall(f"{{{VML}}}shape")
            self.assertEqual(ET.tostring(original_shapes[1]), ET.tostring(updated_shapes[1]))
            anchor = updated_shapes[0].find(f"{{{EXCEL}}}ClientData/{{{EXCEL}}}Anchor")
            self.assertEqual([int(value) for value in anchor.text.split(",")], [1, 67, 0, 2, 4, 0, 5, 2])
        restored = load_workbook(BytesIO(result))
        self.assertEqual(restored.active["D1"].comment.text, "Header note")
        self.assertEqual(restored.active["A2"].comment.text, "Other note")
        self.assertEqual(restored["Other"]["A1"].comment.text, "Other sheet note")

    def test_export_without_error_result_does_not_require_note_drawing(self):
        first = pd.DataFrame({Col.NAME: ["A"], Col.YEAR: [2021], Col.MONTH: [1]})
        data, filename = to_excel_bytes(first, pd.DataFrame(), "result.xlsx")
        workbook = load_workbook(BytesIO(data))
        self.assertEqual(filename, "result.xlsx")
        self.assertIn(SheetName.FIRST_PAYMENT, workbook.sheetnames)
        self.assertNotIn(SheetName.ERROR_RESULT, workbook.sheetnames)


if __name__ == "__main__":
    unittest.main()
