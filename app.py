import streamlit as st
import pandas as pd
import math
from io import BytesIO


def extract_first_payment_month(df):
    """
    최초 납입월 명단을 추출한다.

    Args:
        df (DataFrame): 원본 데이터프레임

    Returns:
        DataFrame: 최초납입월 데이터
    """
    df_first = df[df["코드1"].isin([1, 2, 3]) & (df["코드2"] == 1)].copy()
    df_first = df_first.sort_values(["이름", "해당년", "해당월"], ascending=[True, True, True])
    df_first = df_first.drop_duplicates(subset=["이름"], keep="first")
    df_first = df_first.reset_index(drop=True)
    df_first.index = df_first.index + 1
    return df_first


def detect_errors(df):
    """
    오류를 검출하여 결과를 생성한다.

    Args:
        df (DataFrame): 원본 데이터프레임

    Returns:
        DataFrame: 오류검출 결과
    """
    condition1 = (df["코드1"] == 1) & (df["코드2"].between(1, 7))
    condition2 = (df["코드1"] == 2) & (df["코드2"] == 1)
    condition3 = (df["코드1"] == 3) & (df["코드2"] == 1)

    filtered_df = df[condition1 | condition2 | condition3]
    filtered_df = filtered_df.sort_values(
        ["이름", "해당년", "해당월"], ascending=[False, False, False]
    )

    grouped = filtered_df.groupby(["이름", "해당년", "해당월"])
    valid_groups = []

    for (name, year, month), group in grouped:
        codes = group["코드1"].unique()
        if 1 not in codes:
            continue
        if not (2 in codes or 3 in codes):
            continue

        summed_group = group.groupby("코드1")["입금"].sum().reset_index()
        입금_코드1 = summed_group[summed_group["코드1"] == 1]["입금"].iloc[0]

        summed_group["기준금액"] = None
        summed_group["기준"] = None
        summed_group["구분"] = None

        조건1_만족 = False
        조건2_만족 = False

        if 2 in codes:
            입금_코드2 = summed_group[summed_group["코드1"] == 2]["입금"].iloc[0]
            if (year <= 2018) or (year == 2019 and month <= 3):
                기준값_비율 = 0.3
            else:
                기준값_비율 = 0.4
            기준문구 = f"운영기금 총입금액 {기준값_비율 * 100}%"
            기준값1 = int((입금_코드1 * 기준값_비율) // 1000 * 1000)
            기준값1_1 = int(math.ceil(입금_코드1 * 기준값_비율 / 1000) * 1000)
            if 입금_코드2 < 기준값1:
                조건1_만족 = True
                summed_group.loc[summed_group["코드1"] == 2, "기준금액"] = 기준값1
                summed_group.loc[summed_group["코드1"] == 2, "기준"] = 기준문구
                summed_group.loc[summed_group["코드1"] == 2, "구분"] = "부족"
            elif 입금_코드2 > 기준값1_1:
                조건1_만족 = True
                summed_group.loc[summed_group["코드1"] == 2, "기준금액"] = 기준값1_1
                summed_group.loc[summed_group["코드1"] == 2, "기준"] = 기준문구
                summed_group.loc[summed_group["코드1"] == 2, "구분"] = "초과"

        if 3 in codes:
            입금_코드3 = summed_group[summed_group["코드1"] == 3]["입금"].iloc[0]
            if 입금_코드3 != 입금_코드1:
                조건2_만족 = True
                summed_group.loc[summed_group["코드1"] == 3, "기준금액"] = 입금_코드1
                summed_group.loc[summed_group["코드1"] == 3, "기준"] = "운영기금 총입금액"
                summed_group.loc[summed_group["코드1"] == 3, "구분"] = "불일치"

        if 조건1_만족 or 조건2_만족:
            조건_만족_데이터 = []
            if 조건1_만족:
                조건_만족_데이터.append(summed_group[summed_group["코드1"] == 2])
            if 조건2_만족:
                조건_만족_데이터.append(summed_group[summed_group["코드1"] == 3])
            filtered_summed_group = pd.concat(조건_만족_데이터, ignore_index=True)
            filtered_summed_group["이름"] = name
            filtered_summed_group["해당년"] = year
            filtered_summed_group["해당월"] = month
            valid_groups.append(filtered_summed_group)

    if valid_groups:
        result_df = pd.concat(valid_groups, ignore_index=True)
        result_df = result_df[
            ["이름", "해당년", "해당월", "코드1", "입금", "기준금액", "기준", "구분"]
        ]
        result_df = result_df.sort_values(["이름", "해당년", "해당월", "코드1"])
        result_df = result_df.reset_index(drop=True)
        result_df.index = result_df.index + 1
        return result_df
    else:
        return pd.DataFrame()


def generate_missed_months(df):
    target_codes = [1, 2, 3]
    work = df[["이름", "해당년", "해당월", "코드1", "입금"]].copy()
    work = work.sort_values(["이름", "해당년", "해당월", "코드1"])
    missed_rows = []
    for (name, year, month), group in work.groupby(["이름", "해당년", "해당월"]):
        present_codes = set(group["코드1"].tolist())
        for code in target_codes:
            if code not in present_codes:
                missed_rows.append(
                    {
                        "이름": name,
                        "해당년": year,
                        "해당월": month,
                        "코드1": code,
                        "입금": 0,
                        "기준금액": None,
                        "기준": None,
                        "구분": "미납",
                    }
                )
    if not missed_rows:
        return pd.DataFrame()
    df_missed = pd.DataFrame(missed_rows)
    df_missed = df_missed.sort_values(["이름", "해당년", "해당월", "코드1"])
    return df_missed


def to_excel_bytes(df_first, df_errors, output_filename):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        if not df_first.empty:
            df_first.to_excel(writer, index=True, index_label="번호", sheet_name="최초납입월")
            worksheet_first = writer.sheets["최초납입월"]
            worksheet_first.auto_filter.ref = worksheet_first.dimensions
        if not df_errors.empty:
            df_errors.to_excel(writer, index=True, index_label="번호", sheet_name="오류검출결과")
            worksheet_errors = writer.sheets["오류검출결과"]
            worksheet_errors.column_dimensions["G"].width = 10
            worksheet_errors.column_dimensions["H"].width = 30
            from openpyxl.styles import Alignment

            for row in worksheet_errors.iter_rows(min_row=1, max_row=len(df_errors) + 1):
                for col_idx, cell in enumerate(row):
                    if col_idx == 0:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_idx == 1:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_idx == 2:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_idx == 3:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_idx == 4:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                    elif col_idx == 5:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        if cell.value and str(cell.value).replace(",", "").isdigit():
                            cell.number_format = "#,##0"
                    elif col_idx == 6:
                        cell.alignment = Alignment(horizontal="right", vertical="center")
                        if cell.value and str(cell.value).replace(",", "").isdigit():
                            cell.number_format = "#,##0"
                    elif col_idx == 7:
                        cell.alignment = Alignment(horizontal="left", vertical="center")
                    elif col_idx == 8:
                        cell.alignment = Alignment(horizontal="center", vertical="center")
    buffer.seek(0)
    return buffer.getvalue(), output_filename


st.set_page_config(page_title="인별납부내역 오류검출", layout="wide")
st.title("인별납부내역 오류검출 웹앱")
st.caption("엑셀 파일을 업로드하면 최초납입월과 오류검출 결과를 생성합니다.")

with st.sidebar:
    st.header("입력 파일")
    main_file = st.file_uploader("※정리본 엑셀 (raw 시트 포함)", type=["xlsx"], key="main")
    grad_file = st.file_uploader("오늘공동체졸업생명단.xlsx", type=["xlsx"], key="grad")
    st.divider()
    sheet_name = st.text_input("원본 시트명", value="raw")
    header_row = st.number_input("헤더 시작 행 (0-index)", min_value=0, value=1, step=1)
    run_btn = st.button("처리 실행")

if run_btn:
    if not main_file or not grad_file:
        st.error("두 개의 엑셀 파일을 모두 업로드하세요.")
    else:
        try:
            df = pd.read_excel(main_file, sheet_name=sheet_name, header=header_row)
        except Exception as e:
            st.error(f"원본 엑셀 읽기 오류: {e}")
            st.stop()

        required_cols = ["이름", "해당년", "해당월", "코드1", "코드2", "입금"]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.error(f"누락된 열: {missing}")
            st.stop()

        try:
            df_grad = pd.read_excel(grad_file, sheet_name="졸업생명단", header=None)
            graduation_names = set(df_grad.iloc[:, 0].dropna().astype(str).tolist())
            df = df[df["이름"].astype(str).isin(graduation_names)].copy()
        except Exception as e:
            st.error(f"졸업생명단 처리 오류: {e}")
            st.stop()

        numeric_columns = ["해당년", "해당월", "코드1", "코드2", "입금"]
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
                df[col] = df[col].round().astype("Int64")

        with st.spinner("오류검출 실행 중..."):
            df_errors = detect_errors(df)

        with st.spinner("최초납입월/미납월 생성 중..."):
            df_first_payment = extract_first_payment_month(df[df["코드2"] == 1])
            df_missed = generate_missed_months(
                df[(df["코드2"] == 1) & (df["코드1"].isin([1, 2, 3]))]
            )
            if not df_missed.empty:
                df_errors = pd.concat([df_errors, df_missed], ignore_index=True)
            if not df_errors.empty:
                df_errors = df_errors.sort_values(
                    ["이름", "해당년", "해당월", "코드1"]
                ).reset_index(drop=True)
                df_errors.index = df_errors.index + 1

        st.success("처리가 완료되었습니다.")

        col1, col2 = st.columns(2)
        with col1:
            st.subheader("최초납입월")
            st.dataframe(df_first_payment, use_container_width=True)
        with col2:
            st.subheader("오류검출결과")
            st.dataframe(df_errors, use_container_width=True)

        base_name = getattr(main_file, "name", "result.xlsx")
        output_name = base_name.replace(".xlsx", "_오류검출.xlsx")
        data, out_name = to_excel_bytes(df_first_payment, df_errors, output_name)
        st.download_button(
            label="결과 엑셀 다운로드",
            data=data,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
