import importlib
import unittest
from unittest.mock import patch

import pandas as pd
from nicegui import ui
from nicegui.client import Client

from constants import Col, Status

with patch.object(ui, "run"):
    view = importlib.import_module("app")


class TestResultView(unittest.IsolatedAsyncioTestCase):
    async def test_status_totals_and_people_share_columns_below_status_rows(self):
        errors = pd.DataFrame([
            {Col.NAME: "A", Col.YEAR: 2026, Col.MONTH: 8, Col.STATUS: Status.UNPAID, Col.ERROR_TARGET: "O"},
            {Col.NAME: "A", Col.YEAR: 2026, Col.MONTH: 8, Col.STATUS: Status.INSUFFICIENT, Col.ERROR_TARGET: "O"},
            {Col.NAME: "B", Col.YEAR: 2026, Col.MONTH: 9, Col.STATUS: Status.EXCESS, Col.ERROR_TARGET: "O"},
        ])
        # 저장된 이전 결과에서도 다운로드 파일명으로 기준월을 복원한다.
        summary = {"duration": 1.2, "period_count": 100, "filtered_count": 0, "period": "2026년"}
        client = Client(ui.page("/_result_view_test"))
        self.addCleanup(client.delete)
        with client:
            with ui.column() as parent:
                view.build_result_view(errors, pd.DataFrame(), pd.DataFrame(),
                                       "정리본_26.11.4_오류검출.xlsx", summary)

        table = next(node for node in parent.descendants() if node.tag == "table")
        rows = [
            [node.text for node in row.descendants() if isinstance(node, ui.label)]
            for row in table.descendants() if row.tag == "tr"
        ]
        self.assertEqual(rows, [
            ["구분", "납부 오류", "3개월 경과 오류"],
            ["미납", "1건", "1건"],
            ["부족", "1건", "1건"],
            ["초과", "1건", "0건"],
            ["합계", "3건", "2건"],
            ["대상자", "2명", "1명"],
        ])


if __name__ == "__main__":
    unittest.main()
