import unittest
from unittest.mock import patch

import requests

from data_processor import get_google_sheets_title


def response_with(content):
    response = requests.Response()
    response.status_code = 200
    response.encoding = "utf-8"
    response._content = content.encode("utf-8")
    return response


class TestSheetConnection(unittest.TestCase):
    def test_google_sheet_title_is_preserved(self):
        response = response_with("<html><title>졸업생 원본 - Google Sheets</title></html>")
        with patch("data_processor.requests.get", return_value=response) as get:
            title = get_google_sheets_title("https://example.invalid/sheet/edit")

        self.assertEqual(title, "졸업생 원본")
        get.assert_called_once_with("https://example.invalid/sheet/view", timeout=10)

    def test_direct_csv_with_required_columns_is_ready(self):
        response = response_with("이름,구분\n홍길동,졸업생\n")
        url = "https://example.invalid/export?format=csv"
        with patch("data_processor.requests.get", return_value=response) as get:
            self.assertEqual(get_google_sheets_title(url), "졸업생 명단")
        get.assert_called_once_with(url, timeout=10)

    def test_csv_missing_required_column_is_not_ready(self):
        for content in ["이름\n홍길동\n", "구분\n졸업생\n", "<html>접근 불가</html>"]:
            with self.subTest(content=content):
                with patch("data_processor.requests.get", return_value=response_with(content)):
                    self.assertIsNone(get_google_sheets_title("https://example.invalid/list.csv"))

    def test_network_and_http_errors_are_not_ready(self):
        response = response_with("이름,구분\n홍길동,졸업생\n")
        response.status_code = 403
        with patch("data_processor.requests.get", return_value=response):
            self.assertIsNone(get_google_sheets_title("https://example.invalid/list.csv"))
        with patch("data_processor.requests.get", side_effect=requests.Timeout):
            self.assertIsNone(get_google_sheets_title("https://example.invalid/list.csv"))


if __name__ == "__main__":
    unittest.main()
