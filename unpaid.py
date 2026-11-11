import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.units import pixels_to_points

from constants import Col, SheetName, Status
from excess_shortfall import FUNDS, MONTH_KEYS, TARGET_DIGIT_WIDTH_PX, TARGET_DISPLAY_DPI


UNPAID_FUNDS = "미납 기금명"


def build_unpaid_summary(df_errors):
    columns = [*MONTH_KEYS, UNPAID_FUNDS]
    if df_errors is None or df_errors.empty:
        return pd.DataFrame(columns=columns)
    targets = df_errors.loc[
        df_errors[Col.ERROR_TARGET].eq("O").fillna(False)
        & df_errors[Col.STATUS].eq(Status.UNPAID)
    ]
    if targets.empty:
        return pd.DataFrame(columns=columns)
    summary = targets.groupby(MONTH_KEYS)[Col.FUND_NAME].agg(
        lambda names: ", ".join(fund for fund, _ in FUNDS if fund in set(names))
    ).reset_index(name=UNPAID_FUNDS)
    return summary.sort_values(MONTH_KEYS, ascending=[True, False, False]).reset_index(drop=True)


def write_unpaid_sheet(writer, df_errors):
    summary = build_unpaid_summary(df_errors)
    ws = writer.book.create_sheet(SheetName.UNPAID)
    ws.sheet_view.showGridLines = False
    ws.sheet_view.zoomScale = 100
    ws.sheet_view.zoomScaleNormal = 100
    height = pixels_to_points(36, dpi=TARGET_DISPLAY_DPI)
    ws.sheet_format.defaultRowHeight = height
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    edge = Side(style="thin", color="FF000000")
    border = Border(left=edge, right=edge, top=edge, bottom=edge)
    fills = ["FFFFFFA7", "FFD4E3F8", "FFF8E3D4"]
    for column, pixels in {"A": 110, "B": 110, "C": 220}.items():
        ws.column_dimensions[column].width = pixels / TARGET_DIGIT_WIDTH_PX
    for column, label in enumerate(["이름", "년월", UNPAID_FUNDS], start=1):
        cell = ws.cell(1, column, label)
        cell.fill = PatternFill("solid", fgColor="FFE7E6E6")
        cell.font = Font(size=11, bold=True, color="FF000000")
        cell.alignment = center
        cell.border = border
    ws.row_dimensions[1].height = height
    ws.freeze_panes = "A2"
    if summary.empty:
        ws.merge_cells("A2:C2")
        ws["A2"] = "미납 항목이 없습니다."
        ws["A2"].alignment = center
        ws["A2"].font = Font(size=11)
        ws.row_dimensions[2].height = height
        return

    row_number = 2
    for name, person in summary.groupby(Col.NAME, sort=False):
        first_row = row_number
        for _, record in person.iterrows():
            values = [str(name), f"{int(record[Col.YEAR])}-{int(record[Col.MONTH]):02d}", record[UNPAID_FUNDS]]
            for column, value in enumerate(values, start=1):
                cell = ws.cell(row_number, column, value)
                cell.data_type = "s"
                cell.fill = PatternFill("solid", fgColor=fills[column - 1])
                cell.font = Font(size=11, color="FF000000")
                cell.alignment = center
                cell.border = border
            ws.row_dimensions[row_number].height = height
            row_number += 1
        if row_number - first_row > 1:
            ws.merge_cells(start_row=first_row, start_column=1, end_row=row_number - 1, end_column=1)
        ws.row_dimensions[row_number].height = pixels_to_points(28, dpi=TARGET_DISPLAY_DPI)
        row_number += 1
