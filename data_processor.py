import pandas as pd
import math
from io import BytesIO
import requests
import re
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side
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

def extract_date_from_filename(filename: str) -> tuple[int, int] | None:
    """
    파일명에서 년월을 추출한다.
    패턴: ※정리본_YY.M.D.xlsx (예: ※정리본_26.2.14.xlsx → (2026, 2))
    """
    match = re.search(r"_(\d{2})\.(\d{1,2})\.\d{1,2}\.xlsx?$", filename)
    if not match:
        return None

    yy = int(match.group(1))
    month = int(match.group(2))

    year = 2000 + yy if yy <= 30 else 1900 + yy

    return (year, month)

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
            std = target_amounts.round().astype(int)
        else:
            ratios = pd.Series(1.0, index=work_df.index)
            std = op_deposit

        formula_series = "운영기금 총액 × " + (ratios * 100).astype(int).astype(str) + "%"

        # 부족 검사
        insufficient_mask = ((config["deposit"] < std) & config["has_mask"]).fillna(False)
        insufficient_mask = insufficient_mask & ((std - config["deposit"]) >= 1000)
        if insufficient_mask.any():
            res = work_df[insufficient_mask].copy()
            res[Col.FUND_NAME] = config["name"]
            res[Col.CODE] = config["code"]
            res[Col.STATUS] = Status.INSUFFICIENT
            res[Col.STANDARD] = std[insufficient_mask]
            res[Col.FORMULA] = formula_series[insufficient_mask]
            res[Col.DEPOSIT] = config["deposit"][insufficient_mask]
            res[Col.DIFF] = res[Col.DEPOSIT] - res[Col.STANDARD]
            results.append(res)

        # 초과 검사
        excess_mask = ((config["deposit"] > std) & config["has_mask"]).fillna(False)
        excess_mask = excess_mask & ((config["deposit"] - std) >= 1000)
        if excess_mask.any():
            res = work_df[excess_mask].copy()
            res[Col.FUND_NAME] = config["name"]
            res[Col.CODE] = config["code"]
            res[Col.STATUS] = Status.EXCESS
            res[Col.STANDARD] = std[excess_mask]
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

def generate_missed_months(df: pd.DataFrame, df_first: pd.DataFrame | None = None, filename: str | None = None) -> tuple[pd.DataFrame, int]:
    """
    미납월(누락된 코드)을 생성한다.
    - 코드 1: 11~17 중 하나라도 있으면 납부로 인정. 모두 없으면 미납.
    - 코드 2: 21이 없으면 미납.
    - 코드 3: 31이 없으면 미납.
    - 단, df_first(최초납부월 정보)가 주어지면, 해당 인원의 최초납부월 이전 데이터는 미납에서 제외한다.
    - filename이 주어지면, 기준년월-3개월 이전의 미납에 "과거 미납분"을 비고에 표기한다.
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
    
    ref_date = extract_date_from_filename(filename) if filename else None
    if ref_date:
        cutoff_serial = ref_date[0] * 12 + ref_date[1] - 3
        if "serial" not in missed.columns:
            missed["serial"] = missed[Col.YEAR] * 12 + missed[Col.MONTH]
        missed[Col.REMARKS] = missed.apply(
            lambda row: "과거 미납분" if row["serial"] < cutoff_serial else "",
            axis=1
        )
        missed = missed.drop(columns=["serial"])
    else:
        missed[Col.REMARKS] = ""
    
    missed = missed[RESULT_COLUMNS].sort_values([Col.NAME, Col.YEAR, Col.MONTH, Col.CODE])
    
    return missed, filtered_count

def to_excel_bytes(df_first, df_errors, output_filename, df_summary=None):
    """
    데이터프레임을 엑셀 바이트로 변환 (리팩토링 버전)
    """
    buffer = BytesIO()
    
    def prepare_sheet(df, exclude_cols=None):
        if df is None or df.empty:
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
            ws.column_dimensions["H"].width = 12  # 기준금액
            ws.column_dimensions["I"].width = 26  # 기준금액 산출법
            ws.column_dimensions["J"].width = 9   # 상태
            ws.column_dimensions["K"].width = 12  # 차액
            ws.column_dimensions["L"].width = 41  # 비고
            
            # 오토필터 및 정렬/포맷팅
            last_col = get_column_letter(len(df_e.columns) + 1)
            ws.auto_filter.ref = f"A1:{last_col}{len(df_e) + 1}"

            # 헤더 스타일 설정 (중앙 정렬)
            for cell in ws[1]:
                cell.alignment = Alignment(horizontal="center")

            # 데이터 행 정렬 및 포맷팅 (2행부터)
            # 고대비 노르딕 브루탈리즘 색상 팔레트 정의 (Updated - Darker & Warmer)
            fills = {
                1: PatternFill(start_color="FFFDE047", end_color="FFFDE047", fill_type="solid"), # 이름 (Yellow 300 - Vibrant Yellow)
                6: PatternFill(start_color="FFFDE047", end_color="FFFDE047", fill_type="solid"), # 납부금액 (Yellow 300)
                7: PatternFill(start_color="FFE9D5FF", end_color="FFE9D5FF", fill_type="solid"), # 기준금액 (Purple 200)
                8: PatternFill(start_color="FFF3E8FF", end_color="FFF3E8FF", fill_type="solid"), # 기준금액 산출법 (Purple 100)
                # 차액 배경색 제거 (idx=10)
            }

            # 테두리 스타일 정의
            thin_border = Border(
                left=Side(style='thin', color='FF9CA3AF'),
                right=Side(style='thin', color='FF9CA3AF'),
                top=Side(style='thin', color='FF9CA3AF'),
                bottom=Side(style='thin', color='FF9CA3AF')
            )
            
            thick_bottom_border = Border(
                left=Side(style='thin', color='FF9CA3AF'),
                right=Side(style='thin', color='FF9CA3AF'),
                top=Side(style='thin', color='FF9CA3AF'),
                bottom=Side(style='thick', color='FF9CA3AF')  # 3배 두께
            )
            
            # 헤더에 배경색 및 두꺼운 하단 테두리 적용
            for cell in ws[1]:
                col_idx_zero = cell.col_idx - 1
                if col_idx_zero in fills:
                    cell.fill = fills[col_idx_zero]
                cell.border = thick_bottom_border

            # 동일 그룹 식별 (이름/납부년/납부월) - 그룹 인덱스 추적
            group_row_colors = {}  # {excel_row: color_index} - 홀수/짝수 그룹 구분
            
            if not df_e.empty:
                groups = []
                current_group_start = 0
                current_key = None
                
                for idx, (i, row_data) in enumerate(df_e.iterrows()):
                    key = (row_data[Col.NAME], row_data[Col.YEAR], row_data[Col.MONTH])
                    
                    if current_key is None:
                        current_key = key
                        current_group_start = idx
                    elif key != current_key:
                        groups.append((current_group_start, idx - 1, current_key))
                        current_key = key
                        current_group_start = idx
                
                if current_key is not None:
                    groups.append((current_group_start, len(df_e) - 1, current_key))
                
                # 2개 이상 행을 가진 그룹에 대해 홀수/짝수 인덱스 할당
                group_counter = 0
                for start_idx, end_idx, key in groups:
                    if end_idx > start_idx:  # 2개 이상 행을 가진 그룹만
                        color_index = group_counter % 2  # 0(짝수) 또는 1(홀수)
                        for row_idx in range(start_idx, end_idx + 1):
                            excel_row = row_idx + 2  # 엑셀 행 번호
                            group_row_colors[excel_row] = color_index
                        group_counter += 1

            # 데이터 행 스타일 적용
            red_font = Font(color='FFDC2626', bold=True)  # 진한 붉은색 (홀수 그룹)
            blue_font = Font(color='FF0066FF', bold=True)  # 밝고 쨍한 파란색 (짝수 그룹)
            
            for row in ws.iter_rows(min_row=2, max_row=len(df_e) + 1):
                excel_row_num = row[0].row
                color_index = group_row_colors.get(excel_row_num)  # None, 0(짝수), 1(홀수)
                
                # 현재 행의 기금명 값을 먼저 확인 (코드 컬럼 배경색 동기화용)
                fund_name_cell = ws.cell(row=excel_row_num, column=5)  # 기금명은 5번째 컬럼 (E열)
                fund_value = str(fund_name_cell.value).strip() if fund_name_cell.value else ""
                
                for idx, cell in enumerate(row):
                    # 기본 배경색 적용
                    if idx in fills:
                        cell.fill = fills[idx]
                    
                    # 기본 테두리 적용
                    cell.border = thin_border
                    
                    # 그룹 행의 번호/이름/납부년/납부월에 교대로 색상 적용
                    if color_index is not None and idx in [0, 1, 2, 3]:  # 번호, 이름, 납부년, 납부월
                        if color_index == 1:  # 홀수 그룹
                            cell.font = red_font
                        else:  # 짝수 그룹
                            cell.font = blue_font

                    # [신규] 기금명 컬럼(idx=4)에 배경색 적용
                    if idx == 4:  # 기금명
                        if fund_value == "운영":
                            cell.fill = PatternFill(start_color="FF69A2B0", end_color="FF69A2B0", fill_type="solid")
                        elif fund_value == "협력":
                            cell.fill = PatternFill(start_color="FFFFCAB1", end_color="FFFFCAB1", fill_type="solid")
                        elif fund_value == "복지":
                            cell.fill = PatternFill(start_color="FF659157", end_color="FF659157", fill_type="solid")
                    
                    # [신규] 코드 컬럼(idx=5)에 기금명과 동일한 배경색 적용
                    if idx == 5:  # 코드
                        if fund_value == "운영":
                            cell.fill = PatternFill(start_color="FF69A2B0", end_color="FF69A2B0", fill_type="solid")
                        elif fund_value == "협력":
                            cell.fill = PatternFill(start_color="FFFFCAB1", end_color="FFFFCAB1", fill_type="solid")
                        elif fund_value == "복지":
                            cell.fill = PatternFill(start_color="FF659157", end_color="FF659157", fill_type="solid")
                    
                    # [신규] 상태 컬럼(idx=9)에 텍스트 색상 및 아이콘 적용
                    if idx == 9:  # 상태 (컬럼 순서 변경 후)
                        status_value = str(cell.value).strip() if cell.value else ""
                        if status_value == "미납":
                            cell.font = Font(color='FF9CA3AF', bold=True)  # 연한 회색 (Gray 400)
                            cell.value = "— " + status_value  # em dash 아이콘
                        elif status_value == "부족":
                            cell.font = Font(color='FFEF4444', bold=True)  # 빨간색 (Red 500)
                            cell.value = "↓ " + status_value  # 아래 화살표
                        elif status_value == "초과":
                            cell.font = Font(color='FF059669', bold=True)  # 진한 녹색 (Emerald 600)
                            cell.value = "↑ " + status_value  # 위 화살표
                    
                    # [신규] 차액 컬럼(idx=10)에 상태와 동일한 텍스트 색상 적용 및 양수에 + 기호 추가
                    if idx == 10:  # 차액
                        # 현재 행의 상태 값 확인 (상태는 이제 10번째 컬럼, J열)
                        status_cell = ws.cell(row=excel_row_num, column=10)  # 상태는 10번째 컬럼 (J열)
                        status_value = str(status_cell.value).strip() if status_cell.value else ""
                        
                        # 상태에 따라 텍스트 색상 적용
                        if "미납" in status_value:
                            cell.font = Font(color='FF9CA3AF', bold=True)  # 연한 회색
                        elif "부족" in status_value:
                            cell.font = Font(color='FFEF4444', bold=True)  # 빨간색
                        elif "초과" in status_value:
                            cell.font = Font(color='FF059669', bold=True)  # 진한 녹색
                        
                        # 양수 값에 + 기호 추가
                        if cell.value is not None and isinstance(cell.value, (int, float)) and cell.value > 0:
                            cell.number_format = '+#,##0;-#,##0;0'

                    # 정렬 및 포맷
                    if idx <= 5:  # 번호 ~ 코드
                        cell.alignment = Alignment(horizontal="center")
                    elif idx in [6, 7]:  # 납부금액, 기준금액
                        cell.alignment = Alignment(horizontal="right")
                        if cell.value is not None:
                            cell.number_format = "#,##0"
                    elif idx == 10:  # 차액 (위에서 이미 포맷 설정)
                        cell.alignment = Alignment(horizontal="right")
                    elif idx == 9:  # 상태
                        cell.alignment = Alignment(horizontal="center")
                    elif idx == 8:  # 기준금액 산출법
                        cell.alignment = Alignment(horizontal="left")
                    else:
                        cell.alignment = Alignment(horizontal="center")
            
            # 헤더 행 고정 (2행부터 스크롤)
            ws.freeze_panes = "A2"

            # 헤더 행 고정 (2행부터 스크롤)
            ws.freeze_panes = "A2"

        # 2. 오류요약 시트 (이름별 요약)
        df_s = prepare_sheet(df_summary)
        if df_s is not None:
            df_s.to_excel(writer, index=True, index_label="번호", sheet_name=SheetName.ERROR_SUMMARY)
            ws_s = writer.sheets[SheetName.ERROR_SUMMARY]
            
            # 컬럼 너비 설정
            ws_s.column_dimensions["A"].width = 8   # 번호
            ws_s.column_dimensions["B"].width = 12  # 이름
            ws_s.column_dimensions["C"].width = 9   # 미납
            ws_s.column_dimensions["D"].width = 9   # 부족
            ws_s.column_dimensions["E"].width = 9   # 초과
            ws_s.column_dimensions["F"].width = 12  # 오류건수 합계
            
            # 오토필터
            last_col = get_column_letter(len(df_s.columns) + 1)
            ws_s.auto_filter.ref = f"A1:{last_col}{len(df_s) + 1}"
            
            # 헤더 스타일 설정 (중앙 정렬)
            for cell in ws_s[1]:
                cell.alignment = Alignment(horizontal="center")
            
            # 데이터 행 정렬 (모든 컬럼 중앙 정렬)
            for row in ws_s.iter_rows(min_row=2, max_row=len(df_s) + 1):
                for cell in row:
                    cell.alignment = Alignment(horizontal="center")
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"
            if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"
            
            # 소계 행 추가
            if not df_s.empty:
                # 소계 행 생성 (번호: 빈, 이름: 빈, 미납/부족/초과/합계: 합계)
                # 소계 행 생성 (번호: "소계", 이름: 개수, 미납/부족/초과/합계: 합계)
                total_row_data = {
                    "번호": "",
                    "이름": len(df_s),  # 이름 데이터의 개수
                    "미납": df_s["미납"].sum(),
                    "부족": df_s["부족"].sum(),
                    "초과": df_s["초과"].sum(),
                    "합계": df_s["합계"].sum()
                }
                # 마지막 행 다음에 소계 행 추가
                last_data_row = len(df_s) + 1  # +1은 헤더 고려
                for col_idx, col_name in enumerate(df_s.columns, start=1):
                    cell = ws_s.cell(row=last_data_row + 1, column=col_idx)
                    if col_name in total_row_data:
                        cell.value = total_row_data[col_name]
                    else:
                        cell.value = ""
                
                # 소계 행의 첫 번째 컬럼에 "소계" 텍스트 설정 (번호열 대신)
                ws_s.cell(row=last_data_row + 1, column=1).value = "소계"
                
                # 소계 행 스타일 적용 (굵은 폰트, 중앙 정렬, 테두리)
                thin_border = Border(
                    left=Side(style='thin', color='FF9CA3AF'),
                    right=Side(style='thin', color='FF9CA3AF'),
                    top=Side(style='thin', color='FF9CA3AF'),
                    bottom=Side(style='thin', color='FF9CA3AF')
                )
                for cell in ws_s[last_data_row + 1]:
                    cell.alignment = Alignment(horizontal="center")
                    cell.font = Font(bold=True)  # 모든 셀 굵게 표시
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"
                    cell.border = thin_border  # 테두리 적용
                for cell in ws_s[last_data_row + 1]:
                    cell.alignment = Alignment(horizontal="center")
                    if isinstance(cell.value, (int, float)):
                        cell.number_format = "#,##0"
                        cell.font = Font(bold=True)
                
                # 소계 행을 헤더 바로 뒤(2행)로 이동
                # 현재 마지막 행에 있는 소계를 삭제하고 2행에 삽입
                if last_data_row + 1 > 2:
                    # 마지막 행의 값 저장
                    subtotal_row_values = [cell.value for cell in ws_s[last_data_row + 1]]
                    # 마지막 행 삭제
                    ws_s.delete_rows(last_data_row + 1, 1)
                    # 2행에 소계 행 삽입
                    ws_s.insert_rows(2, 1)
                    # 2행에 값 복원
                    for col_idx, value in enumerate(subtotal_row_values, start=1):
                        ws_s.cell(row=2, column=col_idx).value = value
                        # 스타일 복원 (굵은 폰트, 중앙 정렬, 테두리)
                        thin_border = Border(
                            left=Side(style='thin', color='FF9CA3AF'),
                            right=Side(style='thin', color='FF9CA3AF'),
                            top=Side(style='thin', color='FF9CA3AF'),
                            bottom=Side(style='thin', color='FF9CA3AF')
                        )
                        cell = ws_s.cell(row=2, column=col_idx)
                        cell.alignment = Alignment(horizontal="center")
                        cell.font = Font(bold=True)  # 모든 셀 굵게 표시
                        if isinstance(value, (int, float)):
                            cell.number_format = "#,##0"
                        cell.border = thin_border  # 테두리 적용
                        cell = ws_s.cell(row=2, column=col_idx)
                        cell.alignment = Alignment(horizontal="center")
                        if isinstance(value, (int, float)):
                            cell.number_format = "#,##0"
                            cell.font = Font(bold=True)
            
            # 소계 행이 있으면 오토필터 범위 조정
            if not df_s.empty:
                last_col = get_column_letter(len(df_s.columns) + 1)
                ws_s.auto_filter.ref = f"A1:{last_col}{len(df_s) + 2}"  # +2는 소계 행 포함
            
            # 틀고정 적용 (1행 헤더만 고정, 2행 소계부터 스크롤)
            ws_s.freeze_panes = "A3"  # A3부터 스크롤하면 1,2행이 고정됨


        # 3. 최초납부월 시트
        df_f = prepare_sheet(df_first, exclude_cols=["구분", "코드3"])
        if df_f is not None:
            df_f.to_excel(writer, index=True, index_label="번호", sheet_name=SheetName.FIRST_PAYMENT)
            ws_f = writer.sheets[SheetName.FIRST_PAYMENT]
            # 헤더 중앙 정렬 추가 (최초납부월 시트)
            for cell in ws_f[1]:
                cell.alignment = Alignment(horizontal="center")
                
            # 오토필터 적용
            last_col = get_column_letter(len(df_f.columns) + 1)
            ws_f.auto_filter.ref = f"A1:{last_col}{len(df_f) + 1}"
            
            # 틀고정 적용 (1행 헤더만 고정)
            ws_f.freeze_panes = "A2"
    buffer.seek(0)
    return buffer.getvalue(), output_filename
