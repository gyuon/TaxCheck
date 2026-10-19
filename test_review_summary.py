import unittest

import pandas as pd

from constants import Col, Status
from data_processor import summarize_review_targets


class TestReviewSummary(unittest.TestCase):
    def test_counts_only_targets_and_deduplicates_people(self):
        errors = pd.DataFrame([
            {Col.NAME: "A", Col.STATUS: Status.INSUFFICIENT, Col.ERROR_TARGET: "O", Col.UNPAID_3_MONTHS: ""},
            {Col.NAME: "A", Col.STATUS: Status.EXCESS, Col.ERROR_TARGET: "O", Col.UNPAID_3_MONTHS: ""},
            {Col.NAME: "B", Col.STATUS: Status.UNPAID, Col.ERROR_TARGET: "O", Col.UNPAID_3_MONTHS: "O"},
            {Col.NAME: "B", Col.STATUS: Status.UNPAID, Col.ERROR_TARGET: "O", Col.UNPAID_3_MONTHS: "O"},
            {Col.NAME: "C", Col.STATUS: Status.INSUFFICIENT, Col.ERROR_TARGET: "", Col.UNPAID_3_MONTHS: ""},
            {Col.NAME: "D", Col.STATUS: Status.EXCESS, Col.ERROR_TARGET: "", Col.UNPAID_3_MONTHS: ""},
        ])
        original = errors.copy(deep=True)

        summary = summarize_review_targets(errors)

        self.assertEqual(summary["target_count"], 4)
        self.assertEqual(summary["person_count"], 2)
        self.assertEqual(summary["long_unpaid_count"], 2)
        self.assertEqual(summary["counts"], {
            Status.UNPAID: 2, Status.INSUFFICIENT: 1, Status.EXCESS: 1,
        })
        self.assertEqual(sum(summary["counts"].values()), summary["target_count"])
        pd.testing.assert_frame_equal(errors, original)

    def test_long_unpaid_uses_existing_flag_only_for_unpaid_targets(self):
        errors = pd.DataFrame([
            {Col.NAME: "A", Col.STATUS: Status.UNPAID, Col.ERROR_TARGET: "O", Col.UNPAID_3_MONTHS: "O"},
            {Col.NAME: "A", Col.STATUS: Status.UNPAID, Col.ERROR_TARGET: "O", Col.UNPAID_3_MONTHS: ""},
            {Col.NAME: "B", Col.STATUS: Status.INSUFFICIENT, Col.ERROR_TARGET: "O", Col.UNPAID_3_MONTHS: "O"},
        ])

        self.assertEqual(summarize_review_targets(errors)["long_unpaid_count"], 1)

    def test_empty_results_have_zero_metrics(self):
        summary = summarize_review_targets(pd.DataFrame())

        self.assertEqual(summary["target_count"], 0)
        self.assertEqual(summary["person_count"], 0)
        self.assertEqual(summary["long_unpaid_count"], 0)
        self.assertTrue(all(value == 0 for value in summary["counts"].values()))

    def test_only_excluded_results_have_zero_metrics(self):
        errors = pd.DataFrame([
            {Col.NAME: "A", Col.STATUS: Status.EXCESS, Col.ERROR_TARGET: ""},
            {Col.NAME: "B", Col.STATUS: Status.INSUFFICIENT, Col.ERROR_TARGET: pd.NA},
        ])

        summary = summarize_review_targets(errors)

        self.assertEqual(summary["target_count"], 0)
        self.assertEqual(summary["person_count"], 0)
        self.assertEqual(summary["long_unpaid_count"], 0)

    def test_legacy_results_use_same_thousand_won_rule(self):
        errors = pd.DataFrame([
            {Col.NAME: "A", Col.STATUS: Status.INSUFFICIENT, Col.DIFF: -1000, "비고": ""},
            {Col.NAME: "B", Col.STATUS: Status.EXCESS, Col.DIFF: 1000, "비고": ""},
            {Col.NAME: "C", Col.STATUS: Status.INSUFFICIENT, Col.DIFF: -999, "비고": ""},
            {Col.NAME: "D", Col.STATUS: Status.EXCESS, Col.DIFF: 999, "비고": ""},
            {Col.NAME: "E", Col.STATUS: Status.UNPAID, Col.DIFF: None, "비고": "3개월 이전 미납분"},
        ])
        original = errors.copy(deep=True)

        summary = summarize_review_targets(errors)

        self.assertEqual(summary["target_count"], 3)
        self.assertEqual(summary["person_count"], 3)
        self.assertEqual(summary["long_unpaid_count"], 1)
        self.assertEqual(summary["counts"], {
            Status.UNPAID: 1, Status.INSUFFICIENT: 1, Status.EXCESS: 1,
        })
        pd.testing.assert_frame_equal(errors, original)


if __name__ == "__main__":
    unittest.main()
