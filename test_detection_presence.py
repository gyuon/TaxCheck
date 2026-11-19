import unittest
from unittest.mock import patch
import pandas as pd
from pandas.core.groupby.generic import SeriesGroupBy

from constants import Col, Status
from data_processor import detect_errors


class TestDetectionPresence(unittest.TestCase):
    def test_zero_and_cancelled_payments_are_present_without_group_apply(self):
        rows = []
        for name, payments in [('zero', [(1, 100000), (2, 0)]),
                               ('cancelled', [(1, 100000), (2, 1000), (2, -1000)]),
                               ('absent', [(1, 100000)]),
                               ('no_operating', [(2, 40000), (3, 100000)])]:
            rows.extend({Col.NAME: name, Col.YEAR: 2026, Col.MONTH: 1,
                         Col.CODE1: code, Col.RAW_DEPOSIT: amount} for code, amount in payments)
        source = pd.DataFrame(rows)
        original = source.copy(deep=True)
        with patch.object(SeriesGroupBy, 'apply', side_effect=AssertionError('group apply')):
            result = detect_errors(source)
        self.assertEqual(set(result[Col.NAME]), {'zero', 'cancelled'})
        self.assertTrue(result[Col.STATUS].eq(Status.INSUFFICIENT).all())
        self.assertEqual(result[Col.DIFF].tolist(), [-40000, -40000])
        pd.testing.assert_frame_equal(source, original)
