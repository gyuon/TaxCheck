
import unittest
import ast
import pandas as pd
import numpy as np
import data_processor
import sys
from pathlib import Path
from constants import Col, Status, FundName, FundCode, RESULT_COLUMNS

# Helper to print colored output or distinct separators
def print_section(title):
    print(f"\n{'='*50}")
    print(f"TEST: {title}")
    print(f"{'='*50}")

def print_result(msg):
    print(f" -> {msg}")

class TestDataProcessorVerbose(unittest.TestCase):
    
    def setUp(self):
        # Sample Data (Same as before)
        self.sample_data = {
            Col.NAME: ["UserA", "UserA", "UserB", "UserB", "UserC", "UserC", "UserD"],
            Col.YEAR: [2021, 2021, 2021, 2021, 2021, 2021, 2021],
            Col.MONTH: [1, 1, 1, 1, 1, 1, 2],
            Col.CODE1: [1, 2, 1, 3, 1, 2, 2],
            Col.CODE2: [1, 1, 1, 1, 1, 1, 1],
            Col.RAW_DEPOSIT: [
                100000, 10000,   # UserA: Insufficient Coop
                100000, 50000,   # UserB: Mismatch Welfare
                100000, 50000,   # UserC: Excess Coop
                10000            # UserD: Missed Operating
            ]
        }
        self.df = pd.DataFrame(self.sample_data)

    def test_verify_columns_and_logic_cases(self):
        print_section("검출 로직 및 컬럼 구조 검증 (6개 케이스)")
        
        errors = data_processor.detect_errors(self.df)
        
        # 1. Column Structure
        self.assertEqual(errors.columns.tolist(), RESULT_COLUMNS)
        print_result("컬럼 구조 확인 완료")
        
        # 2. Case 1: Insufficient (부족)
        row_a = errors[errors[Col.NAME] == "UserA"].iloc[0]
        self.assertEqual(row_a[Col.STATUS], Status.INSUFFICIENT)
        self.assertEqual(row_a[Col.DIFF], -30000)
        self.assertIn("%", str(row_a[Col.FORMULA])) # 퍼센트 기호 확인
        print_result(f"Case 1 (부족): UserA - 상태={row_a[Col.STATUS]}, 차액={row_a[Col.DIFF]}, 산출법={row_a[Col.FORMULA]} (Pass)")

        # 3. Case 2: Welfare Insufficient (복지 부족) - 기존 불일치 케이스
        row_b = errors[errors[Col.NAME] == "UserB"].iloc[0]
        self.assertEqual(row_b[Col.STATUS], Status.INSUFFICIENT) # 불일치 제거 후 부족으로 분류
        self.assertEqual(row_b[Col.DIFF], -50000)
        self.assertEqual(row_b[Col.FORMULA], "운영기금 총액(100,000) × 100%")
        print_result(f"Case 2 (복지 부족): UserB - 상태={row_b[Col.STATUS]}, 차액={row_b[Col.DIFF]}, 산출법={row_b[Col.FORMULA]} (Pass)")
        
        # 4. Case 3: Excess (초과)
        row_c = errors[errors[Col.NAME] == "UserC"].iloc[0]
        self.assertEqual(row_c[Col.STATUS], Status.EXCESS)
        self.assertEqual(row_c[Col.DIFF], 10000)
        print_result(f"Case 3 (초과): UserC - 상태={row_c[Col.STATUS]}, 차액={row_c[Col.DIFF]} (Pass)")

        # Generate Missed Months for UserD
        df_miss = self.df[self.df[Col.NAME] == "UserD"].copy()
        missed, _ = data_processor.generate_missed_months(df_miss)
        
        # 5. Case 4: Missed Operating (운영 미납 -> 코드 11~17)
        op_missed = missed[missed[Col.FUND_NAME] == FundName.OPERATING]
        self.assertFalse(op_missed.empty)
        code_val = op_missed.iloc[0][Col.CODE]
        status_val = op_missed.iloc[0][Col.STATUS]
        self.assertEqual(code_val, FundCode.OPERATING_UNPAID)
        self.assertEqual(status_val, Status.UNPAID)
        print_result(f"Case 4 (운영 미납): UserD - 코드={code_val}, 상태={status_val} (Pass)")
        
        # 6. Case 5: Missed Welfare (복지 미납 -> 코드 31)
        welfare_missed = missed[missed[Col.FUND_NAME] == FundName.WELFARE]
        self.assertFalse(welfare_missed.empty)
        self.assertEqual(welfare_missed.iloc[0][Col.CODE], str(FundCode.WELFARE))
        print_result(f"Case 5 (복지 미납): UserD - 코드={welfare_missed.iloc[0][Col.CODE]} (Pass)")

    def test_extra_functionality(self):
        print_section("추가 기능 및 엣지 케이스 검증")
        
        # Normalization
        df_raw = pd.DataFrame({Col.NAME: ["  Test  "]})
        processed = data_processor.normalize_names(df_raw)
        self.assertEqual(processed.iloc[0][Col.NAME], "Test")
        print_result("이름 공백 제거 정규화 (Pass)")
        
        # Excel Generation
        df_dummy = pd.DataFrame({"A": [1]})
        e_bytes, name = data_processor.to_excel_bytes(df_dummy, pd.DataFrame(), "out.xlsx")
        self.assertTrue(len(e_bytes) > 0)
        print_result("엑셀 생성 기능 (Pass)")

    def test_filename_date_filter(self):
        print_section("파일명 기준년월 미납 필터 검증")

        base_data = {
            Col.NAME: ["TestUser", "TestUser"],
            Col.YEAR: [2026, 2026],
            Col.MONTH: [2, 2],
            Col.CODE1: [1, 1],
            Col.CODE2: [1, 3],
            Col.RAW_DEPOSIT: [100000, 0],
        }
        df = pd.DataFrame(base_data)

        # 1. filename 날짜 == 데이터의 납부년/월 → 미납 제외
        missed, _ = data_processor.generate_missed_months(
            df, filename="※정리본_26.2.14.xlsx"
        )
        self.assertTrue(missed.empty, "filename 날짜와 일치하는 미납이 제거되지 않음")
        print_result("Case 1: filename(26.2) == month(2) → 미납 제외 (Pass)")

        # 2. filename 날짜 ≠ 데이터의 납부년/월 → 미납 유지
        missed, _ = data_processor.generate_missed_months(
            df, filename="※정리본_26.1.14.xlsx"
        )
        self.assertFalse(missed.empty, "filename 날짜와 불일치하는 미납이 제거됨")
        self.assertIn(Status.UNPAID, missed[Col.STATUS].values)
        print_result("Case 2: filename(26.1) ≠ month(2) → 미납 유지 (Pass)")

        # 3. filename=None → 모든 미납 유지
        missed, _ = data_processor.generate_missed_months(df, filename=None)
        self.assertFalse(missed.empty, "filename=None 시 미납이 제거됨")
        print_result("Case 3: filename=None → 미납 유지 (Pass)")

        # 4. detect_errors는 filename 인자가 없으므로 부족/초과에 영향 없음
        errors_data = {
            Col.NAME: ["TestUser", "TestUser"],
            Col.YEAR: [2026, 2026],
            Col.MONTH: [2, 2],
            Col.CODE1: [1, 2],
            Col.CODE2: [1, 1],
            Col.RAW_DEPOSIT: [100000, 50000],
        }
        errors = data_processor.detect_errors(pd.DataFrame(errors_data))
        self.assertFalse(errors.empty, "detect_errors가 부족/초과를 검출하지 못함")
        self.assertIn(Status.EXCESS, errors[Col.STATUS].values)
        print_result("Case 4: detect_errors는 filename 무관 → 초과 정상 검출 (Pass)")

    def test_small_discrepancy_detection(self):
        print_section("1,000원 미만 오차 검출 검증")

        data = {
            Col.NAME: ["UserX", "UserX", "UserY", "UserY"],
            Col.YEAR: [2021, 2021, 2021, 2021],
            Col.MONTH: [1, 1, 1, 1],
            Col.CODE1: [1, 2, 1, 2],
            Col.CODE2: [1, 1, 1, 1],
            Col.RAW_DEPOSIT: [
                100000, 40300,   # UserX: Coop 300원 초과 (기준 40,000)
                100000, 39800,   # UserY: Coop 200원 부족 (기준 40,000)
            ],
        }
        df = pd.DataFrame(data)
        errors = data_processor.detect_errors(df)

        user_x = errors[errors[Col.NAME] == "UserX"]
        self.assertFalse(user_x.empty, "300원 초과가 검출되지 않음")
        self.assertEqual(user_x.iloc[0][Col.STATUS], Status.EXCESS)
        self.assertEqual(user_x.iloc[0][Col.DIFF], 300)
        print_result("Case 1: 300원 초과 검출 (Pass)")

        user_y = errors[errors[Col.NAME] == "UserY"]
        self.assertFalse(user_y.empty, "200원 부족이 검출되지 않음")
        self.assertEqual(user_y.iloc[0][Col.STATUS], Status.INSUFFICIENT)
        self.assertEqual(user_y.iloc[0][Col.DIFF], -200)
        print_result("Case 2: 200원 부족 검출 (Pass)")

    def test_cooperation_standard_rule_boundaries(self):
        print_section("협력기금 기준금액 기간별 산출법 검증")

        data = {
            Col.NAME: [
                "BeforeRule", "BeforeRule",
                "Start30", "Start30",
                "End30", "End30",
                "Start40", "Start40",
            ],
            Col.YEAR: [2014, 2014, 2014, 2014, 2019, 2019, 2019, 2019],
            Col.MONTH: [2, 2, 3, 3, 3, 3, 4, 4],
            Col.CODE1: [1, 2, 1, 2, 1, 2, 1, 2],
            Col.CODE2: [1, 1, 1, 1, 1, 1, 1, 1],
            Col.RAW_DEPOSIT: [
                90000, 20000,   # 기준 30,000 = 90,000 / 3 -> 부족 10,000
                100000, 20000,  # 기준 30,000 = 100,000 * 30% -> 부족 10,000
                100000, 20000,  # 기준 30,000 = 100,000 * 30% -> 부족 10,000
                100000, 30000,  # 기준 40,000 = 100,000 * 40% -> 부족 10,000
            ],
        }
        df = pd.DataFrame(data)

        errors = data_processor.detect_errors(df)

        before = errors[errors[Col.NAME] == "BeforeRule"].iloc[0]
        self.assertEqual(before[Col.STANDARD], 30000)
        self.assertEqual(before[Col.FORMULA], "운영기금 총액(90,000) ÷ 3")
        self.assertEqual(before[Col.STATUS], Status.INSUFFICIENT)
        self.assertEqual(before[Col.DIFF], -10000)

        start30 = errors[errors[Col.NAME] == "Start30"].iloc[0]
        self.assertEqual(start30[Col.STANDARD], 30000)
        self.assertEqual(start30[Col.FORMULA], "운영기금 총액(100,000) × 30%")
        self.assertEqual(start30[Col.STATUS], Status.INSUFFICIENT)
        self.assertEqual(start30[Col.DIFF], -10000)

        end30 = errors[errors[Col.NAME] == "End30"].iloc[0]
        self.assertEqual(end30[Col.STANDARD], 30000)
        self.assertEqual(end30[Col.FORMULA], "운영기금 총액(100,000) × 30%")
        self.assertEqual(end30[Col.STATUS], Status.INSUFFICIENT)
        self.assertEqual(end30[Col.DIFF], -10000)

        start40 = errors[errors[Col.NAME] == "Start40"].iloc[0]
        self.assertEqual(start40[Col.STANDARD], 40000)
        self.assertEqual(start40[Col.FORMULA], "운영기금 총액(100,000) × 40%")
        self.assertEqual(start40[Col.STATUS], Status.INSUFFICIENT)
        self.assertEqual(start40[Col.DIFF], -10000)

        print_result("협력기금 기간별 기준금액 산출법 (Pass)")

    def test_cooperation_missed_month_standard_uses_new_rule(self):
        print_section("협력기금 미납 기준금액 신규 산출법 검증")

        data = {
            Col.NAME: ["MissingCoop", "MissingCoop"],
            Col.YEAR: [2014, 2014],
            Col.MONTH: [2, 2],
            Col.CODE1: [1, 3],
            Col.CODE2: [1, 1],
            Col.RAW_DEPOSIT: [90000, 90000],
        }
        df = pd.DataFrame(data)

        missed, _ = data_processor.generate_missed_months(df)
        coop = missed[missed[Col.FUND_NAME] == FundName.COOPERATION].iloc[0]

        self.assertEqual(coop[Col.STATUS], Status.UNPAID)
        self.assertEqual(coop[Col.STANDARD], 30000)
        self.assertEqual(coop[Col.FORMULA], "운영기금 총액(90,000) ÷ 3")

        print_result("협력기금 미납 기준금액 /3 산출법 (Pass)")

    def test_column_name_consistency(self):
        """app.py에서 참조하는 컬럼명이 data_processor.py의 출력과 일치하는지 검증"""
        print_section("컬럼명 일관성 검증 (app.py ↔ data_processor.py)")
        
        # data_processor.py가 생성하는 컬럼 목록
        errors = data_processor.detect_errors(self.df)
        df_miss = self.df[self.df["이름"] == "UserD"].copy()
        missed, _ = data_processor.generate_missed_months(df_miss)
        
        # app.py에서 참조하는 컬럼명 목록 (하드코딩된 참조들)
        app_referenced_columns = [Col.NAME, Col.YEAR, Col.MONTH, Col.CODE, Col.STATUS]
        
        # 오류 검출 결과에서 컬럼 존재 여부 확인
        for col in app_referenced_columns:
            self.assertIn(col, errors.columns.tolist(), f"detect_errors 결과에 '{col}' 컬럼이 없습니다.")
        print_result(f"detect_errors 출력 컬럼 확인: {app_referenced_columns} (Pass)")
        
        # 미납 결과에서 컬럼 존재 여부 확인
        for col in app_referenced_columns:
            self.assertIn(col, missed.columns.tolist(), f"generate_missed_months 결과에 '{col}' 컬럼이 없습니다.")
        print_result(f"generate_missed_months 출력 컬럼 확인: {app_referenced_columns} (Pass)")
        
        # '상태' 컬럼의 값이 예상된 값들인지 확인 (app.py의 counts dict와 일치)
        expected_status_values = {Status.UNPAID, Status.INSUFFICIENT, Status.EXCESS}
        actual_status_values = set(errors[Col.STATUS].unique()) | set(missed[Col.STATUS].unique())
        self.assertTrue(actual_status_values.issubset(expected_status_values), 
                        f"예상치 못한 상태 값: {actual_status_values - expected_status_values}")
        print_result(f"'{Col.STATUS}' 컬럼 값 검증: {actual_status_values} (Pass)")

    def test_health_endpoint_declares_status_code_200(self):
        print_section("헬스체크 엔드포인트 HTTP 200 명시 검증")

        app_source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
        module = ast.parse(app_source)

        health_function = next(
            node
            for node in module.body
            if isinstance(node, ast.AsyncFunctionDef) and node.name == "health"
        )

        route_decorator = next(
            decorator
            for decorator in health_function.decorator_list
            if isinstance(decorator, ast.Call)
            and isinstance(decorator.func, ast.Attribute)
            and decorator.func.attr == "get"
            and decorator.args
            and isinstance(decorator.args[0], ast.Constant)
            and decorator.args[0].value == "/health"
        )

        status_code_keyword = next(
            (keyword for keyword in route_decorator.keywords if keyword.arg == "status_code"),
            None,
        )

        self.assertIsNotNone(status_code_keyword, "/health 라우트에 status_code=200이 명시되어야 함")
        self.assertIsInstance(status_code_keyword.value, ast.Constant)
        self.assertEqual(status_code_keyword.value.value, 200)
        print_result("/health 라우트 status_code=200 명시 (Pass)")

    def test_nicegui_reload_disabled_on_render(self):
        print_section("Render 환경 reload 비활성화 검증")

        app_source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
        module = ast.parse(app_source)

        ui_run_call = next(
            node.value
            for node in module.body
            if isinstance(node, ast.Expr)
            and isinstance(node.value, ast.Call)
            and isinstance(node.value.func, ast.Attribute)
            and node.value.func.attr == "run"
        )

        reload_keyword = next(
            (keyword for keyword in ui_run_call.keywords if keyword.arg == "reload"),
            None,
        )

        self.assertIsNotNone(reload_keyword, "ui.run에 reload 환경 분기가 명시되어야 함")
        self.assertEqual(ast.unparse(reload_keyword.value), "os.environ.get('RENDER') != 'true'")
        print_result("Render에서는 reload=False, 로컬에서는 reload=True 분기 (Pass)")

if __name__ == '__main__':
    # Run tests with verbosity but relying on our custom prints for detail
    suite = unittest.TestLoader().loadTestsFromTestCase(TestDataProcessorVerbose)
    unittest.TextTestRunner(verbosity=0).run(suite)
