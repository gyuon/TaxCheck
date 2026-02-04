
import unittest
import pandas as pd
import numpy as np
import data_processor
import sys
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
        self.assertEqual(row_b[Col.FORMULA], "운영기금 총액 × 100%") # 문구 확인
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

if __name__ == '__main__':
    # Run tests with verbosity but relying on our custom prints for detail
    suite = unittest.TestLoader().loadTestsFromTestCase(TestDataProcessorVerbose)
    unittest.TextTestRunner(verbosity=0).run(suite)
