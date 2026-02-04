import streamlit as st
import pandas as pd
from datetime import datetime
import time
import streamlit.components.v1 as components
from data_processor import (
    get_google_sheets_title, 
    extract_first_payment_month, 
    detect_errors, 
    generate_missed_months, 
    to_excel_bytes,
    normalize_names
)
from constants import Col, Status

# [Fix] FutureWarning: Downcasting object dtype arrays on .fillna is deprecated.
pd.set_option('future.no_silent_downcasting', True)

@st.cache_data(show_spinner=False)
def build_excel_bytes(df_first, df_errors, output_filename):
    # Cache the generated Excel bytes to avoid recomputation on reruns
    return to_excel_bytes(df_first, df_errors, output_filename)


st.set_page_config(page_title="인별납부내역 오류검출", layout="wide")
st.title("인별납부내역 오류검출 웹앱")
st.caption("엑셀 파일을 업로드하면 최초납부월과 오류검출 결과를 생성합니다.")

if "processing" not in st.session_state:
    st.session_state["processing"] = False
if "previous_gsheet_url" not in st.session_state:
    st.session_state["previous_gsheet_url"] = ""
if "cached_sheet_title" not in st.session_state:
    st.session_state["cached_sheet_title"] = None

# --- Custom CSS for UI Enhancements ---
st.markdown("""
<style>
    /* Global Font Tweak */
    html, body, [class*="css"] {
        font-family: 'Pretendard', 'Malgun Gothic', sans-serif;
    }
    
    /* Button Styling */
    div.stButton > button {
        font-weight: bold;
        border-radius: 8px;
        transition: all 0.2s;
    }
    div.stButton > button:hover {
        transform: scale(1.02);
    }
    
    /* Expander Styling */
    .streamlit-expanderHeader {
        font-weight: bold;
        background-color: #f8f9fa;
        border-radius: 8px;
    }

    /* Custom Table Styling (HTML) */
    table {
        width: 100%;
        border-collapse: collapse;
        font-size: 14px;
        border: 1px solid #eee;
    }
    th {
        background-color: #f8f9fa;
        color: #333;
        font-weight: 700 !important; /* Bold Headers */
        padding: 12px 8px;
        text-align: left; /* Left Alignment */
        border-bottom: 2px solid #ddd;
    }
    td {
        padding: 10px 8px;
        border-bottom: 1px solid #f1f3f5;
        vertical-align: middle;
    }
    tr:hover {
        background-color: #f8f9fa;
    }
    
    /* Chip Styling */
    .chip {
        display: inline-block;
        padding: 4px 10px;
        border-radius: 12px;
        font-weight: 600;
        font-size: 12px;
        line-height: 1;
    }
    .chip-excess { background-color: #ffe3e3; color: #c92a2a; } /* 초과 */
    .chip-insufficient { background-color: #fff3bf; color: #e67700; } /* 부족 */
    .chip-unpaid { background-color: #f1f3f5; color: #495057; } /* 미납 */
    
    .table-container {
        width: 100%;
        margin-bottom: 2rem;
    }
</style>
""", unsafe_allow_html=True)

def render_styled_table(df):
    """
    Render dataframe using st.dataframe for native scrolling and performance.
    Uses pandas styler for formatting.
    """
    if df.empty:
        st.write("데이터가 없습니다.")
        return

    # Create a styler object
    styler = df.style
    
    # Format numbers with commas
    format_dict = {}
    for col in df.columns:
        if col in [Col.RAW_DEPOSIT, Col.DEPOSIT, Col.STANDARD, Col.DIFF]:
            format_dict[col] = "{:,.0f}"
    
    if format_dict:
        styler = styler.format(format_dict, na_rep="")
    
    # Render using st.dataframe which handles its own scrolling
    st.dataframe(styler, width="stretch", height=500)

def reset_app():
    st.session_state["processing"] = False
    st.session_state["settings_expanded"] = True
    st.rerun()

# Processing state affects the expander state
if "settings_expanded" not in st.session_state:
    st.session_state["settings_expanded"] = True
if "uploaded_file" not in st.session_state:
    st.session_state["uploaded_file"] = None

# 상단 설정 영역 (Collapsible Settings)
if st.session_state.get("df_errors") is not None or st.session_state.get("processing"):
    # [New] 사용자의 요청에 따라 페이지 높이를 강제로 늘려주는 CSS 주입
    # 입력 영역이 닫히기 전에 미리 공간을 확보
    st.markdown("""
        <style>
            .main .block-container {
                min-height: 500px !important;
            }
        </style>
    """, unsafe_allow_html=True)

with st.expander("입력 및 설정", expanded=st.session_state["settings_expanded"]):
    # 1. Define Layout Grid First
    col_file, col_url = st.columns([1, 1.5], gap="large")
    st.divider()
    col_filter, col_config, col_action = st.columns([2, 2, 1], gap="medium")

    # 2. Render Non-Blocking Elements First (File, Filters, Config, Button)
    
    # A. File Uploader
    with col_file:
        st.markdown("#### 1. 엑셀 파일 입력")
        if st.session_state["uploaded_file"] is None:
            uploaded = st.file_uploader("※정리본 엑셀 (raw 시트 포함)", type=["xlsx"], key="main_uploader")
            if uploaded is not None:
                st.session_state["uploaded_file"] = uploaded
                st.rerun()
            main_file = None
        else:
            main_file = st.session_state["uploaded_file"]
            file_col1, file_col2 = st.columns([0.85, 0.15])
            with file_col1:
                st.info(f"📄 **{main_file.name}**")
            with file_col2:
                if st.button("❌", help="파일 제거"):
                    st.session_state["uploaded_file"] = None
                    st.rerun()

    # B. Year Filter
    with col_filter:
        st.markdown("##### 📅 년도 필터")
        c1, c2 = st.columns(2)
        current_year = datetime.now().year
        with c1:
            start_year = st.number_input("시작 년도", value=2013, step=1)
        with c2:
            end_year = st.number_input("끝 년도", value=current_year, step=1)
            
    # C. Advanced Settings
    with col_config:
        st.markdown("##### ⚙️ 고급 설정")
        c3, c4 = st.columns(2)
        with c3:
            sheet_name = st.text_input("시트명", value="raw")
        with c4:
            header_row = st.number_input("헤더행", min_value=0, value=1, step=1)
        
        # [Removed] use_first_payment_filter = st.checkbox("최초납부월 이전 미납 제외 기능 사용", value=True)
        use_first_payment_filter = True # Fixed to True as per user request
            
    # D. Run Button
    with col_action:
        st.write("") # Spacer
        st.write("") 
        if st.button("🚀 처리 실행", type="primary", width="stretch", disabled=st.session_state["processing"]):
            st.session_state["processing"] = True
            st.session_state["settings_expanded"] = False # [Fix] 처리가 시작되면 설정창 자동 닫기
            st.rerun()

    # 3. Render Blocking Elements Last (URL Check)
    with col_url:
        st.markdown("#### 2. 졸업생 명단 연결")
        default_url = (
            "https://docs.google.com/spreadsheets/d/"
            "1GRPi_kP7V9YBAmS-jZKpUI9pwPCpeuaXBEGlGLHfL3g/edit?gid=0#gid=0"
        )
        gsheet_url = st.text_input("Google Sheets URL", value=default_url)
        
        # Blocking logic here - won't prevent previous elements from showing
        if gsheet_url and gsheet_url != "":
            if (
                gsheet_url != st.session_state["previous_gsheet_url"]
                or st.session_state["cached_sheet_title"] is None
            ):
                with st.spinner("Wait..."):
                    sheet_title = get_google_sheets_title(gsheet_url)
                    st.session_state["cached_sheet_title"] = sheet_title
                    st.session_state["previous_gsheet_url"] = gsheet_url

            if st.session_state["cached_sheet_title"]:
                st.caption(f"✅ 연결됨: **{st.session_state['cached_sheet_title']}**")
            else:
                st.caption("⚠️ 시트 제목을 불러올 수 없습니다.")

def run_processing(main_file, gsheet_url, sheet_name, header_row, start_year, end_year, use_first_payment_filter):
    # Clear previous results and errors
    if "df_errors" in st.session_state: del st.session_state["df_errors"]
    if "result_summary" in st.session_state: del st.session_state["result_summary"]
    if "error_msg" in st.session_state: del st.session_state["error_msg"]

    start_time = time.time()
    try:
        if not main_file or not gsheet_url:
            st.session_state["error_msg"] = "원본 엑셀과 Google Sheets URL을 입력하세요."
            return
            
        try:
            df = pd.read_excel(main_file, sheet_name=sheet_name, header=header_row)
        except Exception as e:
            st.session_state["error_msg"] = f"원본 엑셀 읽기 오류: {e}"
            return
            
        # 0. 컬럼명 정규화 (사용자 요청 용어로 변경: 해당년->납부년, 해당월->납부월)
        rename_map = {"해당년": Col.YEAR, "해당월": Col.MONTH}
        df = df.rename(columns=rename_map)

        # 1. 필수 컬럼 확인
        required_cols = [Col.NAME, Col.YEAR, Col.MONTH, Col.CODE1, Col.CODE2, Col.RAW_DEPOSIT]
        missing = [c for c in required_cols if c not in df.columns]
        if missing:
            st.session_state["error_msg"] = f"누락된 열: {missing}"
            return
        
        df = normalize_names(df)

        # 2. 숫자형 변환
        numeric_columns = [Col.YEAR, Col.MONTH, Col.CODE1, Col.CODE2, Col.RAW_DEPOSIT]
        for col in numeric_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors="coerce")
                df[col] = df[col].round().astype("Int64")

        # 3. 졸업생 명단 필터링
        try:
            def to_csv_url(url: str) -> str:
                if "export?format=csv" in url: return url
                if "/edit" in url: return url.split("/edit")[0] + "/export?format=csv"
                return url

            csv_url = to_csv_url(gsheet_url)
            df_sheet = pd.read_csv(csv_url)

            df_sheet = normalize_names(df_sheet)
            if "구분" not in df_sheet.columns or "이름" not in df_sheet.columns:
                st.session_state["error_msg"] = "Google Sheet에 '이름'과 '구분' 열이 필요합니다."
                return
            
            graduation_names = set(df_sheet[df_sheet["구분"].str.strip() == "졸업생"]["이름"].tolist())
            df = df[df["이름"].isin(graduation_names)].copy()
            
            total_grad_count = len(df) # 졸업생 전체 납부 건수 (년도 필터링 전)

            if df.empty:
                st.session_state["error_msg"] = "졸업생 명단 필터링 후 남은 데이터가 없습니다."
                return
        except Exception as e:
            st.session_state["error_msg"] = f"졸업생명단 처리 오류: {e}"
            return

        # 4. 최초납부월 추출 (전체 데이터 대상)
        df_first_payment = extract_first_payment_month(df)

        # 5. 년도 필터 적용
        if Col.YEAR in df.columns:
            mask = (df[Col.YEAR] >= start_year) & (df[Col.YEAR] <= end_year)
            mask = mask.fillna(False) 
            df_view = df[mask].copy()
            
            period_count = len(df_view) # 필터링된 기간 내 건수

            if df_view.empty:
                st.session_state["error_msg"] = f"선택한 년도 범위({start_year} ~ {end_year})에 해당하는 데이터가 없습니다."
                return
        else:
            df_view = df.copy()
            period_count = len(df_view)

        # 6. 오류 검출
        df_errors = detect_errors(df_view)

        # 7. 미납월 생성 및 병합
        filter_arg = df_first_payment if use_first_payment_filter else None
        df_missed, filtered_count = generate_missed_months(df_view, filter_arg)

        if not df_missed.empty:
            if df_errors.empty:
                df_errors = df_missed
            else:
                # [Fix] FutureWarning: Ensure dtypes are aligned before concat
                # Use pd.to_numeric to safely handle pd.NA (converts to np.nan)
                for col in [Col.DEPOSIT, Col.STANDARD, Col.DIFF]:
                    if col in df_errors.columns:
                        df_errors[col] = pd.to_numeric(df_errors[col], errors="coerce").astype(float)
                    if col in df_missed.columns:
                        df_missed[col] = pd.to_numeric(df_missed[col], errors="coerce").astype(float)
                
                # Align string columns
                for col in [Col.CODE]:
                    if col in df_errors.columns:
                        df_errors[col] = df_errors[col].astype(str)
                    if col in df_missed.columns:
                        df_missed[col] = df_missed[col].astype(str)

                df_missed_aligned = df_missed.reindex(columns=df_errors.columns)
                df_errors = pd.concat([df_errors, df_missed_aligned], ignore_index=True)
        
        if not df_errors.empty:
            df_errors = df_errors.sort_values([Col.NAME, Col.YEAR, Col.MONTH, Col.CODE]).reset_index(drop=True)
            df_errors.index = df_errors.index + 1

        # [Optimized] Cache Priming: Build excel bytes while spinner is still active.
        # This populates the @st.cache_data for to_excel_bytes so there's no lag when the download button is rendered.
        base_name = getattr(main_file, "name", "result.xlsx")
        download_name = base_name.replace(".xlsx", "_오류검출.xlsx")
        _ = to_excel_bytes(df_errors, df_first_payment, download_name)

        # Persist results
        st.session_state["df_errors"] = df_errors
        st.session_state["df_first_payment"] = df_first_payment
        st.session_state["download_name"] = download_name
        
        # [New] 상세 결과 요약 정보 계산
        counts = {Status.UNPAID: 0, Status.INSUFFICIENT: 0, Status.EXCESS: 0}
        if not df_errors.empty:
            type_counts = df_errors[Col.STATUS].value_counts().to_dict()
            for k in counts.keys():
                counts[k] = int(type_counts.get(k, 0))
        
        st.session_state["result_summary"] = {
            "counts": counts,
            "filtered_count": filtered_count,
            "total_grad_count": total_grad_count,
            "period_count": period_count,
            "duration": time.time() - start_time,
            "period": f"{start_year}년 ~ {end_year}년"
        }
    except Exception as e:
        import traceback
        st.session_state["error_msg"] = f"작업 중 오류 발생: {e}"
        st.session_state["error_detail"] = traceback.format_exc()

# Main logic execution
if st.session_state["processing"]:
    st.divider()
    st.subheader("⏳ 처리 실행 중...")
    
    with st.spinner("데이터 처리 중입니다..."):
        run_processing(main_file, gsheet_url, sheet_name, header_row, start_year, end_year, use_first_payment_filter)
    
    st.session_state["processing"] = False
    st.session_state["settings_expanded"] = False # [Fix] 자동으로 설정 창을 닫아 스크롤 계산 유도
    st.rerun()

# Render persisted results or errors
if "error_msg" in st.session_state:
    st.error(st.session_state["error_msg"])
    if "error_detail" in st.session_state:
        with st.expander("🛠️ 기술적 상세 정보 (개발자용)"):
            st.code(st.session_state["error_detail"])

if "df_errors" in st.session_state and isinstance(st.session_state["df_errors"], pd.DataFrame):
    st.divider()
    
    # 0. 결과 요약 메시지 표시 (세션 스테이트에서 불러옴)
    summary = st.session_state.get("result_summary", {})
    if summary:
        counts = summary.get("counts", {})
        c_miss = counts.get(Status.UNPAID, 0)
        c_under = counts.get(Status.INSUFFICIENT, 0)
        c_over = counts.get(Status.EXCESS, 0)
        c_filt = summary.get("filtered_count", 0)
        
        total_errors = sum(counts.values())
        duration = summary.get("duration", 0)
        period = summary.get("period", "-")
        total_grad = summary.get("total_grad_count", 0)
        period_count = summary.get("period_count", 0)

        # 요약 정보 렌더링
        st.success(f"✅ 처리가 완료되었습니다. (소요 시간: {duration:.2f}초)")
        
        col_sum1, col_sum2, col_sum3 = st.columns(3)
        with col_sum1:
            st.markdown(f"**📅 검출 대상 기간**  \n{period}")
        with col_sum2:
            st.markdown(f"**📊 납부내역 건수**  \n총 {total_grad:,}건 (기간내: {period_count:,}건)")
        with col_sum3:
            st.markdown(f"**🔍 오류 검출 내역**  \n총 {total_errors}건 (미납: {c_miss}, 부족: {c_under}, 초과: {c_over})")
        
        if c_filt > 0:
            st.info(f"💡 **참고:** {c_filt}건의 미납 내역이 '최초 납부월 이전'이라 제외되었습니다.")
        elif total_errors == 0:
            st.info("검출된 오류나 미납 내역이 없습니다.")

    # 상단 헤더 및 다운로드 버튼
    header_col, action_col = st.columns([1, 0.25])
    with header_col:
        st.subheader("분석 결과")
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
            type="primary", # Added color
        )

    # 탭 구성 (오류 검출 결과가 기본)
    tab1, tab2 = st.tabs(["오류 검출 결과", "최초 납부월"])
    
    with tab1:
        render_styled_table(st.session_state["df_errors"])
        
    with tab2:
        render_styled_table(st.session_state.get("df_first_payment", pd.DataFrame()))
