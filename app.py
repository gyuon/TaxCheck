import streamlit as st
import pandas as pd
import math
from io import BytesIO
import requests
import re


def get_google_sheets_title(url):
    """
    구글 시트 URL에서 제목을 추출한다.

    Args:
        url (str): 구글 시트 URL

    Returns:
        str: 시트 제목 또는 None
    """
    try:
        # URL이 공유 링크인지 확인하고 적절히 변환
        if "/edit" in url:
            view_url = url.replace("/edit", "/view")
        else:
            view_url = url

        response = requests.get(view_url, timeout=10)
        response.raise_for_status()

        # HTML에서 제목 추출
        title_match = re.search(r"<title>(.*?)</title>", response.text)
        if title_match:
            title = title_match.group(1)
            # "Google Sheets" 제거
            title = title.replace(" - Google Sheets", "").replace(" - Google スプレッドシート", "")
            return title.strip()

        return None
    except Exception:
        return None


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
                기준_라벨 = {1: "운영기금", 2: "협력기금", 3: "복지기금"}.get(code, None)
                missed_rows.append(
                    {
                        "이름": name,
                        "해당년": year,
                        "해당월": month,
                        "코드1": code,
                        "입금": 0,
                        "기준금액": None,
                        "기준": 기준_라벨,
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


@st.cache_data(show_spinner=False)
def build_excel_bytes(df_first, df_errors, output_filename):
    # Cache the generated Excel bytes to avoid recomputation on reruns
    return to_excel_bytes(df_first, df_errors, output_filename)


st.set_page_config(page_title="인별납부내역 오류검출", layout="wide")
st.title("인별납부내역 오류검출 웹앱")
st.caption("엑셀 파일을 업로드하면 최초납입월과 오류검출 결과를 생성합니다.")

if "processing" not in st.session_state:
    st.session_state["processing"] = False
if "previous_gsheet_url" not in st.session_state:
    st.session_state["previous_gsheet_url"] = ""
if "cached_sheet_title" not in st.session_state:
    st.session_state["cached_sheet_title"] = None

with st.sidebar:
    st.header("입력 파일")
    main_file = st.file_uploader("※정리본 엑셀 (raw 시트 포함)", type=["xlsx"], key="main")

    # 기본 구글 시트 URL
    default_url = (
        "https://docs.google.com/spreadsheets/d/"
        "1GRPi_kP7V9YBAmS-jZKpUI9pwPCpeuaXBEGlGLHfL3g/edit?gid=0#gid=0"
    )
    gsheet_url = st.text_input("졸업생 명단 Google Sheets URL", value=default_url)

    # 구글 시트 제목 표시 (URL이 변경될 때만 업데이트)
    if gsheet_url and gsheet_url != "":
        # URL이 변경되었거나 캐시된 제목이 없는 경우에만 새로 가져오기
        if (
            gsheet_url != st.session_state["previous_gsheet_url"]
            or st.session_state["cached_sheet_title"] is None
        ):
            with st.spinner("구글 시트 제목을 불러오는 중..."):
                sheet_title = get_google_sheets_title(gsheet_url)
                st.session_state["cached_sheet_title"] = sheet_title
                st.session_state["previous_gsheet_url"] = gsheet_url

        # 캐시된 제목 표시
        if st.session_state["cached_sheet_title"]:
            st.success(f"📊 **{st.session_state['cached_sheet_title']}**")
        else:
            st.warning("시트 제목을 불러올 수 없습니다.")

    st.divider()
    with st.expander("고급 설정", expanded=False):
        sheet_name = st.text_input("원본 시트명", value="raw")
        header_row = st.number_input("헤더 시작 행 (0-index)", min_value=0, value=1, step=1)

    run_btn = st.button("처리 실행", disabled=st.session_state["processing"])

if run_btn:
    st.session_state["processing"] = True
    try:
        with st.spinner("처리 실행 중..."):
            if not main_file or not gsheet_url:
                st.error("원본 엑셀과 Google Sheets URL을 입력하세요.")
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
                    # Google Sheets CSV export: assumes the provided link is a normal view URL
                    # Convert to export?format=csv to read via pandas
                    def to_csv_url(url: str) -> str:
                        if "export?format=csv" in url:
                            return url
                        if "/edit" in url:
                            base = url.split("/edit")[0]
                            return base + "/export?format=csv"
                        if "gid=" in url and "docs.google.com" in url:
                            # Some share links already point to export; let pandas try directly
                            return url
                        return url

                    csv_url = to_csv_url(gsheet_url)
                    df_sheet = pd.read_csv(csv_url)

                    # Expecting a sheet "명단(전체)" with columns including "이름" and "구분"
                    # If multiple sheets are present, Google export is per sheet;
                    # user should share that sheet's URL.
                    # Filter rows where 구분 == 졸업생 and collect names
                    if "구분" not in df_sheet.columns or "이름" not in df_sheet.columns:
                        raise ValueError(
                            "Google Sheet에 '명단(전체)' 시트의 '이름'과 '구분' 열이 필요합니다."
                        )
                    grad_df = df_sheet[df_sheet["구분"].astype(str).str.strip() == "졸업생"]
                    graduation_names = set(
                        grad_df["이름"].dropna().astype(str).str.strip().tolist()
                    )
                    df = df[df["이름"].astype(str).str.strip().isin(graduation_names)].copy()
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

                # Persist results for rendering after reruns (e.g., download clicks)
                st.session_state["df_errors"] = df_errors
                st.session_state["df_first_payment"] = df_first_payment
                base_name = getattr(main_file, "name", "result.xlsx")
                st.session_state["download_name"] = base_name.replace(".xlsx", "_오류검출.xlsx")
                st.success("처리가 완료되었습니다.")
    finally:
        st.session_state["processing"] = False

# Render persisted results if available
if "df_errors" in st.session_state and isinstance(st.session_state["df_errors"], pd.DataFrame):
    header_col, action_col = st.columns([1, 0.25])
    with header_col:
        st.subheader("오류검출결과")
    with action_col:
        dl_name = st.session_state.get("download_name", "result_오류검출.xlsx")
        data, out_name = build_excel_bytes(
            st.session_state.get("df_first_payment", pd.DataFrame()),
            st.session_state["df_errors"],
            dl_name,
        )
        st.download_button(
            label="결과 엑셀 다운로드",
            data=data,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width="stretch",
        )

    st.dataframe(st.session_state["df_errors"], width="stretch")

    with st.expander("최초납입월", expanded=False):
        st.dataframe(st.session_state.get("df_first_payment", pd.DataFrame()), width="stretch")
