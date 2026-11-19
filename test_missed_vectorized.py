import unittest
from unittest.mock import patch
import pandas as pd

from constants import Col, Status
from data_processor import generate_missed_months


class TestMissedVectorized(unittest.TestCase):
    def test_first_month_exemption_cutoff_and_unknown_amount_without_row_apply(self):
        source = pd.DataFrame([
            {Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 1, Col.CODE1: 1, Col.RAW_DEPOSIT: 100000},
            {Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 3, Col.CODE1: 3, Col.RAW_DEPOSIT: 80000},
        ])
        first = pd.DataFrame([{Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 2}])
        original = source.copy(deep=True)
        with patch.object(pd.DataFrame, 'apply', side_effect=AssertionError('row apply')):
            result, filtered = generate_missed_months(
                source, first, filename='input_26.6.4.xlsx', graduation_names=['A'],
                start_year=2026, start_month=1, end_year=2026, end_month=6,
                exemption_map={'A': ['협력']},
            )
        self.assertTrue(result[Col.STATUS].eq(Status.UNPAID).all())
        self.assertNotIn('협력', result[Col.FUND_NAME].tolist())
        self.assertNotIn(1, result[Col.MONTH].tolist())
        self.assertNotIn(6, result[Col.MONTH].tolist())
        unknown = result[result[Col.MONTH].eq(2)]
        self.assertTrue(unknown[Col.STANDARD].isna().all())
        self.assertTrue(unknown[Col.DIFF].isna().all())
        self.assertTrue(unknown[Col.UNPAID_3_MONTHS].eq('O').all())
        march = result[result[Col.MONTH].eq(3)].iloc[0]
        self.assertEqual((march[Col.STANDARD], march[Col.DIFF]), (80000, -80000))
        self.assertEqual(filtered, 0)
        pd.testing.assert_frame_equal(source, original)

    def test_existing_months_before_first_are_counted_and_unknown_people_kept(self):
        source = pd.DataFrame([
            {Col.NAME: name, Col.YEAR: 2026, Col.MONTH: 1, Col.CODE1: 1, Col.RAW_DEPOSIT: 100000}
            for name in ['A', 'B']
        ])
        first = pd.DataFrame([{Col.NAME: 'A', Col.YEAR: 2026, Col.MONTH: 2}])
        result, filtered = generate_missed_months(source, first)
        self.assertEqual(filtered, 2)
        self.assertEqual(set(result[Col.NAME]), {'B'})
