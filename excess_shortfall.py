import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.units import pixels_to_points

from constants import Col, FundName, SheetName, Status
from excel_styles import StyleTemplates


FUNDS = [(FundName.OPERATING, 1), (FundName.COOPERATION, 2), (FundName.WELFARE, 3)]
MONTH_KEYS = [Col.NAME, Col.YEAR, Col.MONTH]
# 사용자 확인 환경: Windows 125%, Excel 100%, 관측된 최대 숫자 폭 10px.
TARGET_DISPLAY_DPI = 120
TARGET_DIGIT_WIDTH_PX = 10


def filter_comparison_period(df_errors, reference_date, start_year=None, end_year=None):
    """두 비교 시트의 공통 범위: 분석 기간 중 기준월보다 3개월 이상 이전."""
    if df_errors is None or df_errors.empty:
        return pd.DataFrame() if df_errors is None else df_errors.copy()
    if reference_date is None or not 1 <= reference_date[1] <= 12:
        raise ValueError("원본 파일명에서 기준 년월을 추출할 수 없습니다. 예: 정리본_26.11.4.xlsx")
    cutoff = reference_date[0] * 12 + reference_date[1] - 3
    years = pd.to_numeric(df_errors[Col.YEAR], errors="coerce")
    months = pd.to_numeric(df_errors[Col.MONTH], errors="coerce")
    mask = months.between(1, 12) & (years * 12 + months <= cutoff)
    if start_year is not None:
        mask &= years >= start_year
    if end_year is not None:
        mask &= years <= end_year
    return df_errors.loc[mask.fillna(False)].copy()


def build_excess_shortfall_summary(df_errors, df_raw=None):
    """오류 대상 월별 비교표. 원본은 오류 검출 대상 코드로 필터링된 내역이다."""
    columns = MONTH_KEYS + [f"{fund}_{metric}" for fund, _ in FUNDS for metric in [Col.DEPOSIT, Col.STANDARD]]
    if df_errors is None or df_errors.empty:
        return pd.DataFrame(columns=columns)

    targets = df_errors.loc[
        df_errors[Col.ERROR_TARGET].eq("O").fillna(False)
        & df_errors[Col.STATUS].isin([Status.EXCESS, Status.INSUFFICIENT])
    ]
    if targets.empty:
        return pd.DataFrame(columns=columns)

    raw_columns = MONTH_KEYS + [Col.CODE1, Col.RAW_DEPOSIT]
    has_raw = df_raw is not None and set(raw_columns).issubset(df_raw.columns)
    if has_raw:
        payments = df_raw.groupby(MONTH_KEYS + [Col.CODE1])[Col.RAW_DEPOSIT].sum(min_count=1).to_dict()
    else:
        # 원본 없이 호출한 경우, 오류 데이터에 있는 납부금액만 사용하고 미상 금액은 비운다.
        payments = (
            df_errors.groupby(MONTH_KEYS + [Col.FUND_NAME])[Col.DEPOSIT].first().to_dict()
            if Col.DEPOSIT in df_errors.columns else {}
        )
    standards = (
        targets.groupby(MONTH_KEYS + [Col.FUND_NAME])[Col.STANDARD].first().to_dict()
        if Col.STANDARD in targets.columns else {}
    )

    months = targets[MONTH_KEYS].drop_duplicates().sort_values(MONTH_KEYS, ascending=[True, False, False])
    rows = []
    for name, year, month in months.itertuples(index=False, name=None):
        key = (name, year, month)
        row = dict(zip(MONTH_KEYS, key))
        for fund, code in FUNDS:
            row[f"{fund}_{Col.DEPOSIT}"] = payments.get((*key, code), 0) if has_raw else payments.get((*key, fund))
            row[f"{fund}_{Col.STANDARD}"] = standards.get((*key, fund))
        rows.append(row)
    return pd.DataFrame(rows, columns=columns)


def write_excess_shortfall_sheet(writer, df_errors, df_raw=None):
    summary = build_excess_shortfall_summary(df_errors, df_raw)
    ws = writer.book.create_sheet(SheetName.EXCESS_SHORTFALL)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.zoomScaleNormal = 100
    row_height = pixels_to_points(36, dpi=TARGET_DISPLAY_DPI)
    ws.sheet_format.defaultRowHeight = row_height
    centered = Alignment(horizontal="center", vertical="center")
    for column, pixels in {"A": 110, "B": 110, "C": 120, "D": 120, "E": 120}.items():
        # OOXML 너비에는 여백이 이미 포함되므로 UI 문자 수 환산식과 구분한다.
        ws.column_dimensions[column].width = pixels / TARGET_DIGIT_WIDTH_PX

    name_fill = PatternFill("solid", fgColor="FFFFFFA7")
    month_fill = PatternFill("solid", fgColor="FFD4E3F8")
    fund_fill = PatternFill("solid", fgColor="FFF8E3D4")
    edge = Side(style="thin", color="FF000000")
    border = Border(left=edge, right=edge, top=edge, bottom=edge)
    upper_border = Border(left=edge, right=edge, top=edge, bottom=Side())
    lower_border = Border(left=edge, right=edge, top=Side(), bottom=edge)
    regular = Font(size=11, color="FF000000")
    standard_font = Font(size=11, color="FFFF0000")
    bottom_aligned = Alignment(horizontal="center", vertical="bottom", wrap_text=True)
    top_aligned = Alignment(horizontal="center", vertical="top", wrap_text=True)
    styles = StyleTemplates()

    for column, label in enumerate(["이름", "년월", *(fund for fund, _ in FUNDS)], start=1):
        cell = ws.cell(1, column, label)
        styles.apply(cell, 'header', border=border, font=regular, alignment=centered, fill=fund_fill)
    ws.row_dimensions[1].height = row_height
    ws.freeze_panes = "A2"
    if summary.empty:
        ws.merge_cells("A2:E2")
        ws["A2"] = "초과·부족 오류 대상이 없습니다."
        ws["A2"].font = regular
        ws["A2"].alignment = centered
        ws.row_dimensions[2].height = row_height
        return

    row_number = 2
    for name, person in summary.groupby(Col.NAME, sort=False):
        for _, month in person.iterrows():
            # 월별 두 행은 숫자 셀로 유지하되 내부 가로선 없이 한 블록으로 보이게 한다.
            for offset, cells in enumerate(ws.iter_rows(min_row=row_number, max_row=row_number + 1, max_col=5)):
                for cell in cells:
                    components = dict(
                        border=upper_border if offset == 0 else lower_border,
                        font=regular,
                        alignment=bottom_aligned if offset == 0 else top_aligned,
                    )
                    if cell.column == 1:
                        components['fill'] = name_fill
                    elif cell.column == 2:
                        components['fill'] = month_fill
                    styles.apply(cell, ('body', cell.column, offset), **components)
            ws.row_dimensions[row_number].height = pixels_to_points(34, dpi=TARGET_DISPLAY_DPI)
            ws.row_dimensions[row_number + 1].height = row_height
            name_cell = ws.cell(row_number, 1, str(name))
            name_cell.data_type = "s"
            ws.cell(row_number, 2, f"{int(month[Col.YEAR])}-{int(month[Col.MONTH]):02d}")

            for column, (fund, _) in enumerate(FUNDS, start=3):
                for offset, metric in [(0, Col.DEPOSIT), (1, Col.STANDARD)]:
                    value = month[f"{fund}_{metric}"]
                    cell = ws.cell(row_number + offset, column, None if pd.isna(value) else value)
                    components = {'number_format': '#,##0'}
                    if metric == Col.STANDARD:
                        components['font'] = standard_font
                    styles.apply(cell, metric, **components)
            row_number += 2

        ws.row_dimensions[row_number].height = pixels_to_points(28, dpi=TARGET_DISPLAY_DPI)
        row_number += 1
