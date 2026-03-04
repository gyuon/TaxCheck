"""
컬럼명 및 상태값 상수 정의
data_processor.py, app.py, test_suite.py에서 공통으로 사용
"""

class Col:
    """오류 검출 결과 시트 컬럼명"""
    NAME = "이름"
    YEAR = "납부년"
    MONTH = "납부월"
    FUND_NAME = "기금명"
    CODE = "코드"
    DEPOSIT = "납부금액"
    STATUS = "상태"
    STANDARD = "기준금액"
    FORMULA = "기준금액 산출법"
    DIFF = "차액"
    REMARKS = "비고"
    
    # 원본 데이터 컬럼 (입력용)
    CODE1 = "코드1"
    CODE2 = "코드2"
    RAW_DEPOSIT = "입금"


class Status:
    """상태 컬럼 값"""
    UNPAID = "미납"
    INSUFFICIENT = "부족"
    EXCESS = "초과"
    
    @classmethod
    def all_values(cls):
        return {cls.UNPAID, cls.INSUFFICIENT, cls.EXCESS}


class SheetName:
    """엑셀 시트 이름"""
    ERROR_RESULT = "오류검출결과"
    ERROR_SUMMARY = "오류요약"
    FIRST_PAYMENT = "최초납부월"


class FundName:
    """기금명 컬럼 값"""
    OPERATING = "운영"
    COOPERATION = "협력"
    WELFARE = "복지"


class FundCode:
    """기금 코드"""
    OPERATING = 11
    OPERATING_UNPAID = "11~17"
    COOPERATION = 21
    WELFARE = 31


# 결과 시트의 컬럼 순서
RESULT_COLUMNS = [
    Col.NAME, Col.YEAR, Col.MONTH, Col.FUND_NAME, Col.CODE,
    Col.DEPOSIT, Col.STANDARD, Col.FORMULA, Col.STATUS, Col.DIFF, Col.REMARKS
]
