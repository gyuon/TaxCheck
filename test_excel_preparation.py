import unittest

from excel_preparation import ExcelPreparation, ExcelPreparationStatus


class TestExcelPreparation(unittest.TestCase):
    def test_completed_file_becomes_downloadable(self):
        preparation = ExcelPreparation()

        generation = preparation.start()

        self.assertEqual(preparation.status, ExcelPreparationStatus.GENERATING)
        self.assertEqual(preparation.presentation.text, "결과 파일 준비 중")
        self.assertEqual(preparation.presentation.icon, "sym_o_overview")
        self.assertTrue(preparation.presentation.disabled)
        self.assertTrue(preparation.presentation.pulsing)

        self.assertTrue(preparation.complete(generation, b"xlsx"))
        self.assertEqual(preparation.status, ExcelPreparationStatus.READY)
        self.assertEqual(preparation.data, b"xlsx")
        self.assertEqual(preparation.presentation.text, "결과 파일 다운로드")
        self.assertEqual(preparation.presentation.icon, "download")
        self.assertFalse(preparation.presentation.disabled)
        self.assertFalse(preparation.presentation.pulsing)

    def test_failed_file_can_be_retried(self):
        preparation = ExcelPreparation()
        generation = preparation.start()

        self.assertTrue(preparation.fail(generation, "write failed"))
        self.assertEqual(preparation.status, ExcelPreparationStatus.FAILED)
        self.assertEqual(preparation.error, "write failed")
        self.assertEqual(preparation.presentation.text, "파일 준비 다시 시도")
        self.assertEqual(preparation.presentation.icon, "refresh")
        self.assertFalse(preparation.presentation.disabled)

        retry_generation = preparation.start()
        self.assertGreater(retry_generation, generation)
        self.assertEqual(preparation.status, ExcelPreparationStatus.GENERATING)
        self.assertIsNone(preparation.error)
        self.assertIsNone(preparation.data)

    def test_stale_generation_cannot_replace_current_result(self):
        preparation = ExcelPreparation()
        stale_generation = preparation.start()
        current_generation = preparation.start()

        self.assertFalse(preparation.complete(stale_generation, b"stale"))
        self.assertFalse(preparation.fail(stale_generation, "stale failure"))
        self.assertEqual(preparation.status, ExcelPreparationStatus.GENERATING)
        self.assertIsNone(preparation.data)
        self.assertIsNone(preparation.error)

        self.assertTrue(preparation.complete(current_generation, b"current"))
        self.assertEqual(preparation.data, b"current")

    def test_invalidated_generation_result_is_discarded(self):
        preparation = ExcelPreparation()
        generation = preparation.start()

        preparation.invalidate()

        self.assertFalse(preparation.complete(generation, b"stale"))
        self.assertEqual(preparation.status, ExcelPreparationStatus.IDLE)
        self.assertIsNone(preparation.data)
        self.assertIsNone(preparation.error)


if __name__ == "__main__":
    unittest.main()
