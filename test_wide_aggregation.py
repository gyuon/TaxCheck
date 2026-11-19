import unittest
from unittest.mock import patch

import pandas as pd

from constants import Col, Status
from wide_summary import build_wide_error_summary_by_person, build_wide_error_summary_by_fund


class TestWideAggregation(unittest.TestCase):
    def test_status_filtering_does_not_repeat_for_every_person_and_year(self):
        errors = pd.DataFrame([
            {Col.NAME: f'P{person}', Col.YEAR: year, Col.STATUS: Status.EXCESS, Col.DIFF: 1000, Col.STANDARD: 2000}
            for person in range(5) for year in range(2013, 2027)
        ])
        original = pd.DataFrame.__getitem__
        calls = []

        def counted(frame, key):
            if isinstance(key, str) and key == Col.STATUS:
                calls.append(key)
            return original(frame, key)

        with patch.object(pd.DataFrame, '__getitem__', counted):
            summary = build_wide_error_summary_by_person(errors, 2013, 2026)
        self.assertLessEqual(len(calls), 20)
        self.assertEqual(summary['총계_초과건수'].tolist(), [14] * 5)
        self.assertEqual(summary['총계_초과금액'].tolist(), [14000] * 5)

    def test_missing_group_outside_period_and_unknown_standard(self):
        errors = pd.DataFrame([
            {Col.NAME: None, Col.FUND_NAME: '운영', Col.YEAR: 2026, Col.STATUS: Status.UNPAID, Col.STANDARD: None, Col.DIFF: None},
            {Col.NAME: 'B', Col.FUND_NAME: '협력', Col.YEAR: 2025, Col.STATUS: Status.EXCESS, Col.STANDARD: 1000, Col.DIFF: 9000},
            {Col.NAME: 'A', Col.FUND_NAME: '복지', Col.YEAR: 2026, Col.STATUS: Status.INSUFFICIENT, Col.STANDARD: 1000, Col.DIFF: -3000},
        ])
        original = errors.copy(deep=True)
        result = build_wide_error_summary_by_person(errors, 2026, 2026)
        self.assertEqual(result.loc[result[Col.NAME].eq('B'), '총계_총건수'].item(), 0)
        self.assertEqual(result.loc[result[Col.NAME].isna(), '총계_금액모름'].item(), 1)
        self.assertEqual(result.iloc[0][Col.NAME], 'A')
        funds = build_wide_error_summary_by_fund(errors, 2026, 2026)
        self.assertEqual(funds[Col.FUND_NAME].tolist(), ['운영', '협력', '복지'])
        self.assertEqual(funds['총계_총건수'].tolist(), [1, 0, 1])
        pd.testing.assert_frame_equal(errors, original)

    def test_total_rounding_uses_unrounded_amounts_across_years(self):
        errors = pd.DataFrame([
            {Col.NAME: 'A', Col.YEAR: year, Col.STATUS: Status.EXCESS, Col.STANDARD: 1, Col.DIFF: 0.6}
            for year in [2025, 2026]
        ])
        result = build_wide_error_summary_by_person(errors, 2025, 2026).iloc[0]
        self.assertEqual(result['2025_초과금액'], 0)
        self.assertEqual(result['2026_초과금액'], 0)
        self.assertEqual(result['총계_초과금액'], 1)

    def test_no_matching_period_still_keeps_people_with_zero_totals(self):
        errors = pd.DataFrame([{Col.NAME: 'A', Col.YEAR: 2013, Col.STATUS: Status.UNPAID, Col.STANDARD: None}])
        for start, end in [(2026, 2026), (2026, 2025)]:
            with self.subTest(start=start, end=end):
                result = build_wide_error_summary_by_person(errors, start, end)
                self.assertEqual(result[Col.NAME].tolist(), ['A'])
                self.assertEqual(result['총계_총건수'].tolist(), [0])
                self.assertEqual(result['총계_금액모름'].tolist(), [0])


if __name__ == '__main__':
    unittest.main()
