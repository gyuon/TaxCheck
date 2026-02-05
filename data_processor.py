import pandas as pd
import math
from io import BytesIO
import requests
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment
import numpy as np
from constants import Col, Status, FundName, FundCode, RESULT_COLUMNS, SheetName

def normalize_names(df: pd.DataFrame) -> pd.DataFrame:
    """
    이름 열의 공백을 제거하는 공통 헬퍼 함수.
    """
    if Col.NAME in df.columns:
        df[Col.NAME] = df[Col.NAME].astype(str).str.strip()
    return df

def get_google_sheets_title(url: str) -> str | None:
    """
    구글 시트 URL에서 제목을 추출한다.
    """
    try:
        if "/edit" in url:
            view_url = url.replace("/edit", "/view")
        else:
            view_url = url

        response = requests.get(view_url, timeout=10)
        response.raise_for_status()

        title_match = re.search(r"<title>(.*?)</title>", response.text)
        if title_match:
            title = title_match.group(1)
            title = title.replace(" - Google Sheets", "").replace(" - Google スプレッドシート", "")
            return title.strip()

        return None
    except Exception:
        return None

def extract_first_payment_month(df: pd.DataFrame) -> pd.DataFrame:
    """
    최초 납부월 명단을 추출한다.
    조건:
      - (코드1==1 & 코드2 1~7) 
      - (코드1==2 & 코드2==1) 
      - (코드1==3 & 코드2==1)
    """
    # 0. 데이터 전처리 (이름 공백 제거)
    df_work = normalize_names(df.copy())
        
    # 1. 조건 필터링
    condition = (
        ((df_work[Col.CODE1] == 1) & (df_work[Col.CODE2].between(1, 7))) |
        ((df_work[Col.CODE1].isin([2, 3])) & (df_work[Col.CODE2] == 1))
    ).fillna(False)
    df_first = df_work[condition].copy()
    
    # 2. 불필요한 열 제거
    exclude_cols = ["구분", "코드3"]
    valid_cols = [c for c in df_first.columns if not str(c).startswith("Unnamed") and c not in exclude_cols]
    df_first = df_first[valid_cols]
    
    # 3. 정렬 및 중복 제거 (가장 과거 날짜 유지)
    df_first = df_first.sort_values([Col.NAME, Col.YEAR, Col.MONTH], ascending=[True, True, True])
    df_first = df_first.drop_duplicates(subset=[Col.NAME], keep="first")
    df_first = df_first.reset_index(drop=True)
    df_first.index = df_first.index + 1
    return df_first

def detect_errors(df: pd.DataFrame) -> pd.DataFrame:
    """
    오류를 검출하여 결과를 생성한다. (가속화 버전)
    """
    # 0. 데이터 전처리 (이름 공백 제거)
    target_df = normalize_names(df.copy())

    # 1. 대상 데이터 필터링
    # 조건: (코드1==1 & 코드2 1~7) OR (코드1==2 & 코드2==1) OR (코드1==3 & 코드2==1)
    condition = (
        ((target_df[Col.CODE1] == 1) & (target_df[Col.CODE2].between(1, 7))) |
        ((target_df[Col.CODE1] == 2) & (target_df[Col.CODE2] == 1)) |
        ((target_df[Col.CODE1] == 3) & (target_df[Col.CODE2] == 1))
    ).fillna(False)
    target_df = target_df[condition].copy()
    
    # 2. 피벗 테이블 생성: (이름, 해당년, 해당월) 별로 각 코드(1,2,3)의 입금 합계 계산
    pivot = target_df.pivot_table(
        index=[Col.NAME, Col.YEAR, Col.MONTH], 
        columns=Col.CODE1, 
        values=Col.RAW_DEPOSIT, 
        aggfunc="sum", 
        fill_value=0
    )
    
    # 필요한 컬럼이 없으면 생성
    for col in [1, 2, 3]:
        if col not in pivot.columns:
            pivot[col] = 0
            
    # 작업 편의를 위해 컬럼 이름 변경
    pivot.columns = ["운영_입금", "협력_입금", "복지_입금"]
    pivot = pivot.reset_index()
    
    # 원본 데이터에서 해당 그룹에 어떤 코드들이 존재하는지 확인 (입금액 0원 등 특이 케이스 대응)
    existence = target_df.groupby([Col.NAME, Col.YEAR, Col.MONTH])[Col.CODE1].apply(set)
    valid_indices = [idx for idx, codes in existence.items() if 1 in codes and (2 in codes or 3 in codes)]
    
    if not valid_indices:
        return pd.DataFrame()

    # pivot 필터링
    pivot_idx = pivot.set_index([Col.NAME, Col.YEAR, Col.MONTH])
    valid_mask = pivot_idx.index.isin(valid_indices)
    work_df = pivot_idx[valid_mask].copy()
    
    results = []
    
    # 벡터화 연산을 위해 컬럼 추출
    op_deposit = work_df["운영_입금"]
    years = work_df.index.get_level_values(Col.YEAR)
    months = work_df.index.get_level_values(Col.MONTH)
    
    # 3. 공통 검사 루프 (협력기금 & 복지기금)
    fund_configs = [
        {
            "name": FundName.COOPERATION,
            "code": FundCode.COOPERATION,
            "deposit": work_df["협력_입금"],
            "has_mask": pd.Series([2 in c for c in existence[work_df.index]], index=work_df.index),
            "type": "coop"
        },
        {
            "name": FundName.WELFARE,
            "code": FundCode.WELFARE,
            "deposit": work_df["복지_입금"],
            "has_mask": pd.Series([3 in c for c in existence[work_df.index]], index=work_df.index),
            "type": "welf"
        }
    ]

    for config in fund_configs:
        # 기금별 기준금액 및 산출법 설정
        if config["type"] == "coop":
            # 협력기금 비율 결정 (2019년 3월 이전 0.3, 이후 0.4)
            ratio_mask = (years <= 2018) | ((years == 2019) & (months <= 3))
            ratios = pd.Series(np.where(ratio_mask, 0.3, 0.4), index=work_df.index)
            
            target_amounts = op_deposit * ratios
            # 기준값: 부족(내림), 초과(올림)
            std_min = (target_amounts // 1000 * 1000).astype(int)
            std_max = (np.ceil(target_amounts / 1000) * 1000).astype(int)
        else:
            # 복지기금: 운영기금과 1:1 동일해야 함
            ratios = pd.Series(1.0, index=work_df.index) # 100%
            std_min = op_deposit
            std_max = op_deposit

        # [최적화] 공통 수식 생성 (× 기호 및 % 형식 적용)
        formula_series = "운영기금 총액 × " + (ratios * 100).astype(int).astype(str) + "%"

        # 부족 검사
        insufficient_mask = ((config["deposit"] < std_min) & config["has_mask"]).fillna(False)
        if insufficient_mask.any():
            res = work_df[insufficient_mask].copy()
            res[Col.FUND_NAME] = config["name"]
            res[Col.CODE] = config["code"]
            res[Col.STATUS] = Status.INSUFFICIENT
            res[Col.STANDARD] = std_min[insufficient_mask]
            res[Col.FORMULA] = formula_series[insufficient_mask]
            res[Col.DEPOSIT] = config["deposit"][insufficient_mask]
            res[Col.DIFF] = res[Col.DEPOSIT] - res[Col.STANDARD]
            results.append(res)

        # 초과 검사
        excess_mask = ((config["deposit"] > std_max) & config["has_mask"]).fillna(False)
        if excess_mask.any():
            res = work_df[excess_mask].copy()
            res[Col.FUND_NAME] = config["name"]
            res[Col.CODE] = config["code"]
            res[Col.STATUS] = Status.EXCESS
            res[Col.STANDARD] = std_max[excess_mask]
            res[Col.FORMULA] = formula_series[excess_mask]
            res[Col.DEPOSIT] = config["deposit"][excess_mask]
            res[Col.DIFF] = res[Col.DEPOSIT] - res[Col.STANDARD]
            results.append(res)
        
    if not results:
        return pd.DataFrame()
        
    final_df = pd.concat(results).reset_index()
    final_df[Col.REMARKS] = ""
    
    # 컬럼 정리
    final_df = final_df[RESULT_COLUMNS]
    
    # 코드 컬럼을 문자열로 변환 (pyarrow 호환성)
    final_df[Col.CODE] = final_df[Col.CODE].astype(str)
    
    # 정렬: 이름 -> 해당년 -> 해당월 -> 기금명(코드 순)
    final_df = final_df.sort_values([Col.NAME, Col.YEAR, Col.MONTH, Col.CODE])
    
    final_df = final_df.reset_index(drop=True)
    final_df.index = final_df.index + 1
    
    return final_df

def generate_missed_months(df: pd.DataFrame, df_first: pd.DataFrame = None) -> tuple[pd.DataFrame, int]:
    """
    미납월(누락된 코드)을 생성한다.
    - 코드 1: 11~17 중 하나라도 있으면 납부로 인정. 모두 없으면 미납.
    - 코드 2: 21이 없으면 미납.
    - 코드 3: 31이 없으면 미납.
    - 단, df_first(최초납부월 정보)가 주어지면, 해당 인원의 최초납부월 이전 데이터는 미납에서 제외한다.
    """
    # 1. 대상 데이터 필터링 및 "대표 코드(CanonicalCode)" 부여
    # (코드1==1 & 코드2 1~7) -> 1
    # (코드1==2 & 코드2==1)   -> 2
    # (코드1==3 & 코드2==1)   -> 3
    
    # 먼저 전체 데이터 중 유효한 코드 범위를 가진 것만 남김 (또는 계산을 위해 플래그 생성)
    # 복사를 떠서 작업
    target_df = normalize_names(df.copy())
    
    # Canonical Code 매핑
    target_df["canonical"] = 0
    mask_op = ((target_df[Col.CODE1] == 1) & (target_df[Col.CODE2].between(1, 7))).fillna(False)
    mask_coop = ((target_df[Col.CODE1] == 2) & (target_df[Col.CODE2] == 1)).fillna(False)
    mask_welf = ((target_df[Col.CODE1] == 3) & (target_df[Col.CODE2] == 1)).fillna(False)
    
    target_df.loc[mask_op, "canonical"] = 1
    target_df.loc[mask_coop, "canonical"] = 2
    target_df.loc[mask_welf, "canonical"] = 3
    
    # 유효한 canonical 코드를 가진 행만 남김 (이름/년/월 그룹 파악용)
    # 하지만 "미납"을 판단하려면, 해당 월에 '다른 납부내역'은 있는데 특정 코드가 빠진 것을 찾아야 함.
    # 즉, 그룹핑 기준은 "유효한 납부내역이 하나라도 있는 (이름, 해당년, 해당월)"
    valid_rows = target_df[target_df["canonical"] > 0].copy()
    
    # [수정] 그룹핑 기준 변경: 유효 코드가 없더라도, 어쨌든 데이터에 존재하는 월이면 미납 검사 대상이 되어야 함.
    # 따라서 groups는 target_df 전체에서 추출
    groups = target_df[[Col.NAME, Col.YEAR, Col.MONTH]].drop_duplicates()
    
    # 3. 각 그룹별로 필요한 Canonical Code (1, 2, 3)를 Cartesian Product로 생성
    groups["key"] = 1
    codes = pd.DataFrame({"canonical": [1, 2, 3], "key": [1, 1, 1]})
    
    expected = pd.merge(groups, codes, on="key").drop("key", axis=1)
    
    # 4. 실제 존재하는 Canonical Code 확인
    # "납부했다"고 인정받으려면 valid_rows(canonical > 0)에 있어야 함.
    # 따라서 existing은 valid_rows에서 추출
    existing = valid_rows[[Col.NAME, Col.YEAR, Col.MONTH, "canonical"]].drop_duplicates()
    existing["present"] = True
    
    # 5. 병합 및 미납 필터링
    merged = pd.merge(expected, existing, on=[Col.NAME, Col.YEAR, Col.MONTH, "canonical"], how="left")
    missed = merged[merged["present"].isna()].copy()
    
    if missed.empty:
        return pd.DataFrame(), 0

    # [신규] 최초 납부월 이전 데이터 필터링
    filtered_count = 0 
    if df_first is not None and not df_first.empty:
        # 1. df_first에서 (이름) -> (Year * 12 + Month) 매핑 생성
        first_map = df_first.set_index(Col.NAME)[[Col.YEAR, Col.MONTH]]
        first_map["serial"] = first_map[Col.YEAR] * 12 + first_map[Col.MONTH]
        first_serial_dict = first_map["serial"].to_dict()
        
        # 2. missed에 serial 컬럼 추가
        missed["serial"] = missed[Col.YEAR] * 12 + missed[Col.MONTH]
        
        # 3. 필터링 함수
        def is_after_first(row):
            try:
                name = row[Col.NAME]
                if name not in first_serial_dict:
                    return True # 최초 납부 정보 없으면 유지
                return row["serial"] >= first_serial_dict[name]
            except Exception as e:
                # 에러 발생 시 데이터 컨텍스트 포함
                raise ValueError(f"데이터 검증 중 오류 발생 (이름: {row.get(Col.NAME,'?')}, 년월: {row.get(Col.YEAR,'?')}-{row.get(Col.MONTH,'?')}): {e}")
        
        # 마스크 계산
        # [수정] .fillna 시 object 타입 배열의 다운캐스팅 FutureWarning 해결
        mask_series = missed.apply(is_after_first, axis=1)
        # 결과가 object 타입인지 확인 후 결측치를 채우고 bool 타입으로 변환
        mask = mask_series.infer_objects(copy=False).fillna(False).astype(bool)
        # 제외된 건수 계산
        filtered_count = (~mask).sum()
        
        missed = missed[mask]
        
        if missed.empty:
            return pd.DataFrame(), filtered_count
        
    # 6. 결과 포맷팅
    label_map = {1: FundName.OPERATING, 2: FundName.COOPERATION, 3: FundName.WELFARE}
    code_map = {1: FundCode.OPERATING, 2: FundCode.COOPERATION, 3: FundCode.WELFARE}
    
    missed[Col.FUND_NAME] = missed["canonical"].map(label_map)
    # [수정] 호환되지 않는 dtype 설정 시 발생하는 FutureWarning 해결
    # Col.CODE가 처음부터 object/문자열로 처리되도록 함
    missed[Col.CODE] = missed["canonical"].map(code_map).astype(str)
    
    # [수정] 운영기금(1) 미납 시 코드를 "11~17"로 변경
    missed.loc[missed["canonical"] == 1, Col.CODE] = "11~17"

    missed[Col.DEPOSIT] = 0
    missed[Col.STATUS] = Status.UNPAID
    missed[Col.STANDARD] = pd.NA
    missed[Col.FORMULA] = ""
    missed[Col.DIFF] = pd.NA
    missed[Col.REMARKS] = ""
    
    # 필요 없어진 컬럼 제거
    missed = missed[RESULT_COLUMNS].sort_values([Col.NAME, Col.YEAR, Col.MONTH, Col.CODE])
    
    return missed, filtered_count

def to_excel_bytes(df_first, df_errors, output_filename):
    """
    데이터프레임을 엑셀 바이트로 변환 (리팩토링 버전)
    """
    buffer = BytesIO()
    
    def prepare_sheet(df, exclude_cols=None):
        if df.empty:
            return None
        # Unnamed 컬럼 제거
        cols = [c for c in df.columns if not str(c).startswith("Unnamed")]
        if exclude_cols:
            cols = [c for c in cols if c not in exclude_cols]
        return df[cols]

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        # 1. 오류검출결과 시트 (사용자 요청으로 앞에 위치)
        df_e = prepare_sheet(df_errors)
        if df_e is not None:
            df_e.to_excel(writer, index=True, index_label="번호", sheet_name=SheetName.ERROR_RESULT)
            ws = writer.sheets[SheetName.ERROR_RESULT]
            
            # 컬럼 너비 설정 (비율 조정: 80/110 ≈ 0.727)
            # A: 번호 (80px), B: 이름 (100px), C-E: 기본 (90px)
            # F: 코드 (80px), G,I,K: 납부금액/기준금액/차액 (110px)
            # H: 상태 (90px), J: 기준금액 산출법 (250px), L: 비고 (400px)
            ws.column_dimensions["A"].width = 8   # 번호
            ws.column_dimensions["B"].width = 10  # 이름
            ws.column_dimensions["C"].width = 9   # 납부년
            ws.column_dimensions["D"].width = 9   # 납부월
            ws.column_dimensions["E"].width = 9   # 기금명
            ws.column_dimensions["F"].width = 8   # 코드
            ws.column_dimensions["G"].width = 12  # 납부금액
            ws.column_dimensions["H"].width = 9   # 상태
            ws.column_dimensions["I"].width = 12  # 기준금액
            ws.column_dimensions["J"].width = 26  # 기준금액 산출법
            ws.column_dimensions["K"].width = 12  # 차액
            ws.column_dimensions["L"].width = 41  # 비고
            
            # 오토필터 및 정렬/포맷팅
            last_col = get_column_letter(len(df_e.columns) + 1)
            ws.auto_filter.ref = f"A1:{last_col}{len(df_e) + 1}"

            # 헤더 스타일 설정 (중앙 정렬)
            for cell in ws[1]:
                cell.alignment = Alignment(horizontal="center")

            # 데이터 행 정렬 및 포맷팅 (2행부터)
            for row in ws.iter_rows(min_row=2, max_row=len(df_e) + 1):
                for idx, cell in enumerate(row):
                    # 0: index(번호), 1: 이름, 2: 납부년, 3: 납부월, 4: 기금명, 5: 코드
                    # 6: 납부금액, 7: 상태, 8: 기준금액, 9: 기준금액 산출법, 10: 차액, 11: 비고
                    
                    if idx <= 5:  # 번호 ~ 코드
                        cell.alignment = Alignment(horizontal="center")
                    elif idx in [6, 8, 10]:  # 납부금액, 기준금액, 차액
                        cell.alignment = Alignment(horizontal="right")
                        if cell.value is not None:
                            cell.number_format = "#,##0"
                    elif idx == 7:  # 상태
                        cell.alignment = Alignment(horizontal="center")
                    elif idx == 9:  # 기준금액 산출법
                        cell.alignment = Alignment(horizontal="left")
                    else:
                        cell.alignment = Alignment(horizontal="center")

        # 2. 최초납부월 시트
        df_f = prepare_sheet(df_first, exclude_cols=["구분", "코드3"])
        if df_f is not None:
            df_f.to_excel(writer, index=True, index_label="번호", sheet_name=SheetName.FIRST_PAYMENT)
            ws_f = writer.sheets[SheetName.FIRST_PAYMENT]
            # 헤더 중앙 정렬 추가 (최초납부월 시트)
            for cell in ws_f[1]:
                cell.alignment = Alignment(horizontal="center")
                    
    buffer.seek(0)
    return buffer.getvalue(), output_filename
