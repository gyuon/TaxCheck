import streamlit as st
import pandas as pd
from datetime import datetime
from typing import cast
import time
import streamlit.components.v1 as components
import st_aggrid  # noqa: F401
from data_processor import (
    get_google_sheets_title, 
    extract_first_payment_month, 
    detect_errors, 
    generate_missed_months, 
    to_excel_bytes,
    normalize_names
)
from constants import Col, Status

# [수정] .fillna 시 object 타입 배열의 다운캐스팅 FutureWarning 해결
pd.set_option('future.no_silent_downcasting', True)

@st.cache_data(show_spinner=False)
def build_excel_bytes(df_first, df_errors, output_filename, df_summary=None):
    # 재실행 시 재계산을 방지하기 위해 생성된 엑셀 바이트를 캐싱함
    return to_excel_bytes(df_first, df_errors, output_filename, df_summary)


st.set_page_config(page_title="인별납부내역 오류검출", layout="wide")
# --- 노르딕 브루탈리스트 헤더 및 타이틀 ---
st.markdown(f"""
    <div style="display: flex; align-items: center; padding: 0px 40px; height: 80px; border-bottom: 1.5px solid #1F2937; margin-bottom: 24px;">
        <span style="font-size: 22px; font-weight: 900; letter-spacing: 1.5px; color: #1F2937;">TAXCHECK</span>
        <div style="margin-left: 64px; display: flex; gap: 40px;">
            <span style="font-size: 14px; font-weight: bold; color: #1F2937;">홈</span>
            <span style="font-size: 14px; font-weight: medium; color: #9CA3AF;">상세 분석</span>
            <span style="font-size: 14px; font-weight: medium; color: #9CA3AF;">기록 보관</span>
        </div>
        <div style="margin-left: auto; display: flex; align-items: center; gap: 20px;">
            <span style="font-size: 12px; font-weight: 900; color: #1F2937;">시스템 가동 중</span>
            <div style="width: 40px; height: 20px; border: 1.5px solid #1F2937; position: relative;">
                <div style="width: 12px; height: 12px; background-color: #1F2937; position: absolute; left: 4px; top: 3px;"></div>
            </div>
        </div>
    </div>
    <div style="padding: 0px 40px; margin-bottom: 24px;">
        <h1 style="font-size: 42px; font-weight: 700; letter-spacing: -1.5px; color: #1F2937; margin: 0;">데이터 분석 대시보드</h1>
    </div>
""", unsafe_allow_html=True)

if "processing" not in st.session_state:
    st.session_state["processing"] = False
if "previous_gsheet_url" not in st.session_state:
    st.session_state["previous_gsheet_url"] = ""
if "cached_sheet_title" not in st.session_state:
    st.session_state["cached_sheet_title"] = None
# ... (render_styled_table, reset_app 함수는 아래에 정의됨) ...

# --- UI 개선을 위한 노르딕 브루탈리스트 CSS ---
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;700;900&display=swap');

    /* 전역 배경 및 폰트 설정 */
    .main {
        background-color: #FFFFFF !important;
        font-family: 'Inter', sans-serif;
    }
    
    /* 텍스트 색상 고정 */
    h1, h2, h3, h4, p, span, label {
        color: #1F2937 !important;
    }

    /* [중요] 모든 입력창 모서리 완전 직각화 (강력한 셀렉터) */
    div[data-testid="stNumberInput"] *, 
    div[data-testid="stTextInput"] *,
    div[data-baseweb="input"] *,
    div[data-baseweb="base-input"] * {
        border-radius: 0px !important;
    }

    /* 입력창 배경색 및 라인 정돈 */
    div[data-baseweb="input"] {
        border: none !important;
        background-color: #FFFFFF !important;
    }

    /* 노르딕 브루탈리스트 카드 스타일 (졸업생 명단 연결 등) */
    div[data-testid="stColumn"]:has(div.nordic-marker) {
        background-color: #DBEAFE !important;
        padding: 16px 20px !important;
        margin-bottom: 8px !important;
        border-radius: 0px !important;
        display: flex;
        flex-direction: column;
        transition: all 0.2s;
    }
    div[data-testid="stColumn"]:has(div.nordic-marker):hover {
        box-shadow: 10px 10px 0px #1F2937;
        transform: translate(-3px, -3px);
    }
    
    /* 카드 내 제목 스타일 */
    div[data-testid="stColumn"] h5 {
        color: #1F2937 !important;
        font-weight: 900 !important;
        font-size: 1.3rem !important;
        margin-top: 0 !important;
        margin-bottom: 12px !important;
        border-bottom: 2.5px solid #1F2937 !important;
        padding-bottom: 6px !important;
    }

    /* 분석 실행 및 다운로드 버튼 스타일 통합: 고대비 노르딕 브루탈리즘 */
    div.stButton > button, div.stDownloadButton > button {
        background-color: #1F2937 !important;
        color: #FFFFFF !important;
        border: 3px solid #000000 !important;
        border-radius: 0px !important;
        font-weight: 900 !important;
        padding: 12px 0px !important;
        letter-spacing: 1px;
        transition: all 0.1s;
        width: 100% !important;
    }
    div.stButton > button p, div.stButton > button span,
    div.stDownloadButton > button p, div.stDownloadButton > button span {
        color: #FFFFFF !important;
        font-weight: 900 !important;
    }
    div.stButton > button:hover, div.stDownloadButton > button:hover {
        background-color: #374151 !important;
        box-shadow: 6px 6px 0px #9CA3AF;
        transform: translate(-1px, -1px);
    }

    /* 구분선 스타일 */
    hr {
        margin: 4rem 0 !important;
        border-color: #1F2937 !important;
        opacity: 1 !important;
    }

    [data-testid="stFileUploaderDropzone"] {
        border-radius: 0px !important;    /* 테두리 곡률 제거 */
        background-color: #FFF9E6 !important; /* 배경색 흰색으로 변경 */
    }

    /* 1. "Drag and drop file here" 문구 수정 */
    [data-testid="stFileUploaderDropzoneInstructions"] div span:first-child {
        font-size: 0px; /* 기존 텍스트 숨김 */
    }
    [data-testid="stFileUploaderDropzoneInstructions"] div span:first-child::after {
        content: "여기에 파일을 끌어다 놓으세요";
        font-size: 16px; /* Streamlit 기본 폰트 크기 */
        visibility: visible;
        display: block;
    }

    /* 2. "Browse files" 버튼 문구 수정 */
    [data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"] {
        font-size: 0px !important; /* 기존 'Browse files' 숨김 */
        border-radius: 0px !important;
        background-color: #FFFFFF !important;
    }
    [data-testid="stFileUploader"] [data-testid="stBaseButton-secondary"]::after {
        content: "파일 찾기";
        font-size: 14px; /* 버튼 텍스트 크기 */
        visibility: visible;
        display: block;
    }

    /* 3. (선택사항) "Limit 20MB per file..." 문구 수정 */
    [data-testid="stFileUploaderDropzoneInstructions"] div span:last-child {
        font-size: 0px;
    }
    [data-testid="stFileUploaderDropzoneInstructions"] div span:last-child::after {
        content: "최대 20MB • XLSX 파일만 가능";
        font-size: 12px;
        visibility: visible;
        display: block;
        margin-top: 5px;
    }
</style>
""", unsafe_allow_html=True)

def render_styled_table(df):
    """
    노르딕 브루탈리스트 스타일의 커스텀 HTML 테이블을 렌더링함.
    장식을 배제하고 데이터 본연의 구조(그리드)를 드러냄.
    """
    if df.empty:
        st.write("데이터가 없습니다.")
        return

    # 성능을 위해 상위 200건만 렌더링 (브루탈리스트 미학 강조)
    display_df = df.head(200)
    
    # 컬럼 한글화 맵
    col_map = {
        "이름": "이름",
        Col.YEAR: "납부년",
        Col.MONTH: "납부월",
        Col.STATUS: "상태",
        Col.CODE: "코드",
        Col.DEPOSIT: "납부액",
        Col.STANDARD: "기준액",
        Col.DIFF: "차액"
    }
    
    html = '<div style="padding: 0 40px;"><table class="nordic-table">'
    # 헤더
    html += '<thead><tr>'
    for col in display_df.columns:
        label = col_map.get(col, col)
        html += f'<th>{label}</th>'
    html += '</tr></thead>'
    
    # 바디
    html += '<tbody>'
    for _, row in display_df.iterrows():
        html += '<tr>'
        for val in row:
            # 숫자 천단위 콤마 처리
            f_val = f"{val:,.0f}" if isinstance(val, (int, float)) else str(val)
            html += f'<td>{f_val}</td>'
        html += '</tr>'
    html += '</tbody></table>'
    
    if len(df) > 200:
        html += f'<p style="font-size: 11px; color: #9CA3AF; margin-top: 8px;">* Showing first 200 of {len(df)} records. Download Excel for full data.</p>'
    
    html += '</div>'
    st.markdown(html, unsafe_allow_html=True)


def reset_app():
    keys_to_drop = [
        "df_errors", "df_first_payment", "result_summary",
        "error_msg", "error_detail", "run_params", "processing",
        "uploaded_file", "cached_sheet_title", "download_name", "button_clicked"
    ]
    for k in keys_to_drop:
        if k in st.session_state:
            del st.session_state[k]
    st.rerun()

# 처리 상태가 Expander(입력창) 열림 상태에 영향을 줌
if "settings_expanded" not in st.session_state:
    st.session_state["settings_expanded"] = True
if "uploaded_file" not in st.session_state:
    st.session_state["uploaded_file"] = None

# 상단 설정 영역 (접고 펼 수 있는 설정창)
if st.session_state.get("df_errors") is not None or st.session_state.get("processing"):
    # [신규] 입력 영역이 닫히기 전에 미리 공간을 확보하여 스크롤 계산 유도
    st.markdown("""
        <style>
            .main .block-container {
                min-height: 500px !important;
            }
        </style>
    """, unsafe_allow_html=True)

# 상단 설정 영역 (물리적 컬럼 카드 레이아웃)
if "df_errors" not in st.session_state and not st.session_state.get("processing", False):
    st.markdown("<div style='padding: 0 20px;'>", unsafe_allow_html=True) # 컬럼 패딩 보간
    c_file, c_url, c_filter, c_config = st.columns(4, gap="large")

    # 1. 원본 엑셀 업로드
    with c_file:
        st.markdown("<div class='nordic-marker'></div>", unsafe_allow_html=True)
        st.markdown("##### 원본 엑셀 업로드")
        if st.session_state["uploaded_file"] is None:
            uploaded = st.file_uploader(
                "이곳에 파일을 드래그하거나 클릭하세요", 
                type=["xlsx"], 
                key="main_uploader", 
                label_visibility="collapsed"
            )
            if uploaded is not None:
                st.session_state["uploaded_file"] = uploaded
                st.rerun()
        else:
            main_file = st.session_state["uploaded_file"]
            st.success(f"준비됨: {main_file.name}") # st.info -> st.success (긍정적 메시지)
            if st.button("파일 다시 선택", key="reset_file", use_container_width=True): # 버튼 텍스트 변경
                st.session_state["uploaded_file"] = None
                st.rerun()

    # 2. 졸업생 명단 연결
    with c_url:
        st.markdown("<div class='nordic-marker'></div>", unsafe_allow_html=True)
        st.markdown("##### 졸업생 명단 연결")
        default_url = "https://docs.google.com/spreadsheets/d/1GRPi_kP7V9YBAmS-jZKpUI9pwPCpeuaXBEGlGLHfL3g/edit?gid=0#gid=0"
        gsheet_url = st.text_input("URL 입력", value=default_url, label_visibility="collapsed")
        
        if gsheet_url:
            if gsheet_url != st.session_state["previous_gsheet_url"] or st.session_state["cached_sheet_title"] is None:
                with st.spinner("연결 중..."):
                    sheet_title = get_google_sheets_title(gsheet_url)
                    st.session_state["cached_sheet_title"] = sheet_title
                    st.session_state["previous_gsheet_url"] = gsheet_url
            
            if st.session_state["cached_sheet_title"]:
                st.markdown(f"""
                    <div style="display: flex; align-items: center; gap: 8px; padding-top: 8px; color: #1F2937;">
                        <span style="font-size: 18px;">🔗</span>
                        <span style="font-weight: 700; font-size: 14px;">{st.session_state['cached_sheet_title']}</span>
                    </div>
                """, unsafe_allow_html=True)
            else:
                st.markdown("""
                    <div style="display: flex; align-items: center; gap: 8px; padding-top: 8px; color: #EF4444;">
                        <span style="font-size: 14px; font-weight: 600;">⚠️ 연결 실패 (URL 확인)</span>
                    </div>
                """, unsafe_allow_html=True)
        else:
            st.markdown("""
                <div style="display: flex; align-items: center; gap: 8px; padding-top: 8px; color: #9CA3AF;">
                    <span style="font-size: 13px; font-weight: 500;">졸업생 명단 URL을 입력해 주세요.</span>
                </div>
            """, unsafe_allow_html=True)

    # 3. 분석 기간 설정
    with c_filter:
        st.markdown("<div class='nordic-marker'></div>", unsafe_allow_html=True)
        st.markdown("##### 분석 기간 설정")
        cur_year = datetime.now().year
        s_year = st.number_input("시작", value=2013, step=1)
        e_year = st.number_input("종료", value=cur_year, step=1)

    # 4. 고급 설정
    with c_config:
        st.markdown("<div class='nordic-marker'></div>", unsafe_allow_html=True)
        st.markdown("##### 고급 설정")
        s_name = st.text_input("시트명", value="raw")
        h_row = st.number_input("헤더행", min_value=0, value=1, step=1)
        u_filter = True # 고정
        
    st.markdown("</div>", unsafe_allow_html=True)
    
    # 중앙 정렬 분석 실행 버튼
    btn_col1, btn_col2, btn_col3 = st.columns([1, 1, 1])
    with btn_col2:
        if not st.session_state.get("button_clicked", False):
            if st.button("분석 실행", type="primary", use_container_width=True, key="analyze_button", disabled=st.session_state["processing"]):
                # 실행 시점의 파라미터 캡처
                st.session_state["run_params"] = {
                    "main_file": st.session_state.get("uploaded_file"),
                    "gsheet_url": gsheet_url,
                    "sheet_name": s_name,
                    "header_row": h_row,
                    "start_year": s_year,
                    "end_year": e_year,
                    "use_filter": u_filter
                }
                st.session_state["button_clicked"] = True
                st.session_state["processing"] = True
                
                # JavaScript로 버튼 즉시 숨기기
                components.html("""
                    <script>
                        const btn = document.querySelector('button[kc="analyze_button"]');
                        if (btn) btn.style.display = 'none';
                    </script>
                """, height=0)
                
                st.rerun()
def run_processing(main_file, gsheet_url, sheet_name, header_row, start_year, end_year, use_first_payment_filter):
    # 기존 결과 및 오류 메시지 초기화
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
            
            graduation_names = list(set(df_sheet[df_sheet["구분"].str.strip() == "졸업생"]["이름"].tolist()))
            df = df[df["이름"].isin(graduation_names)].copy()
            
            total_grad_count = len(df) # 졸업생 전체 납부 건수 (년도 필터링 전)

            if df.empty:
                st.session_state["error_msg"] = "졸업생 명단 필터링 후 남은 데이터가 없습니다."
                return
        except Exception as e:
            st.session_state["error_msg"] = f"졸업생명단 처리 오류: {e}"
            return

        # 4. 최초납부월 추출 (전체 데이터 대상)
        df_first_payment = extract_first_payment_month(cast(pd.DataFrame, df))

        # 5. 년도 필터 적용
        if Col.YEAR in df.columns:
            mask = (df[Col.YEAR] >= start_year) & (df[Col.YEAR] <= end_year)
            mask = mask.fillna(False) 
            df_view: pd.DataFrame = cast(pd.DataFrame, df[mask].copy())
            
            period_count = len(df_view) 

            if df_view.empty:
                st.session_state["error_msg"] = f"선택한 년도 범위({start_year} ~ {end_year})에 해당하는 데이터가 없습니다."
                return
        else:
            df_view = cast(pd.DataFrame, df.copy())
            period_count = len(df_view)

        # 6. 오류 검출
        df_errors = detect_errors(df_view)

        # 7. 미납월 생성 및 병합
        filter_arg = df_first_payment if use_first_payment_filter else None
        df_missed, filtered_count = generate_missed_months(df_view, filter_arg, filename=main_file.name)

        if not df_missed.empty:
            if df_errors.empty:
                df_errors = df_missed
            else:
                for col in [Col.DEPOSIT, Col.STANDARD, Col.DIFF]:
                    if col in df_errors.columns:
                        df_errors[col] = cast(pd.Series, pd.to_numeric(df_errors[col], errors="coerce")).astype(float)
                    if col in df_missed.columns:
                        df_missed[col] = cast(pd.Series, pd.to_numeric(df_missed[col], errors="coerce")).astype(float)
                
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

        # [최적화] 캐시 준비: 스피너가 활성 상태일 때 엑셀 바이트를 미리 생성함.
        # 이렇게 하면 다운로드 버튼이 렌더링될 때 지연 시간 없이 즉시 다운로드 가능함.
        base_name = getattr(main_file, "name", "result.xlsx")
        download_name = base_name.replace(".xlsx", "_오류검출.xlsx")
        
        # 이름별 오류 요약 생성 (엑셀용) - 캐시 전에 먼저 생성
        df_summary = pd.DataFrame()
        if not df_errors.empty:
            summary_by_name = cast(
                pd.DataFrame,
                df_errors.groupby([Col.NAME, Col.STATUS])
                .size()
                .unstack(fill_value=0)
                .astype(int)
                .reset_index()
            )
            
            for status in [Status.UNPAID, Status.INSUFFICIENT, Status.EXCESS]:
                if status not in summary_by_name.columns:
                    summary_by_name[status] = 0
            
            summary_by_name = cast(pd.DataFrame, summary_by_name[[Col.NAME, Status.UNPAID, Status.INSUFFICIENT, Status.EXCESS]])
            
            summary_by_name["합계"] = summary_by_name[Status.UNPAID] + summary_by_name[Status.INSUFFICIENT] + summary_by_name[Status.EXCESS]
            
            summary_by_name = cast(pd.DataFrame, summary_by_name[summary_by_name["합계"] > 0])
            
            summary_by_name = cast(pd.DataFrame, summary_by_name.sort_values("합계", ascending=False)).reset_index(drop=True)
            summary_by_name.index = summary_by_name.index + 1
            
            # 컬럼명 변경
            df_summary = summary_by_name.rename(columns={
                Col.NAME: "이름",
                Status.UNPAID: "미납",
                Status.INSUFFICIENT: "부족",
                Status.EXCESS: "초과"
            })
        
        _ = to_excel_bytes(df_errors, df_first_payment, download_name, df_summary)

        st.session_state["df_summary"] = df_summary

        # 결과 저장
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

# 메인 로직 실행
if st.session_state.get("processing"):
    st.divider()
    st.subheader("⏳ 처리 실행 중...")
    
    with st.spinner("데이터 처리 중입니다..."):
        # 캡처된 파라미터 사용
        params = st.session_state.get("run_params", {})
        run_processing(
            params.get("main_file"), 
            params.get("gsheet_url"), 
            params.get("sheet_name"), 
            params.get("header_row"), 
            params.get("start_year"), 
            params.get("end_year"), 
            params.get("use_filter")
        )
    
    st.session_state["processing"] = False
    st.session_state["settings_expanded"] = False
    st.rerun()

# 에러 메시지 표시
if "error_msg" in st.session_state:
    st.error(st.session_state["error_msg"])
    if "error_detail" in st.session_state:
        with st.expander("🛠️ 기술적 상세 정보 (개발자용)"):
            st.code(st.session_state["error_detail"])

# 성공 결과 렌더링 (에러 블록과 분리)
if "df_errors" in st.session_state:
    # 결과 영역 스타일
    st.markdown("""
        <style>
            /* 분석 결과 영역 상단 여백 감소 */
            [data-testid="stVerticalBlock"] > [style*="flex-direction: column"] > [data-testid="stVerticalBlock"]:has(h2) {
                padding-top: 0px !important;
                margin-top: 0px !important;
            }
            
            /* 결과 헤더 부분 여백 감소 */
            div[data-testid="stMarkdownContainer"] h2 {
                margin-top: 0px !important;
                padding-top: 8px !important;
            }
        </style>
    """, unsafe_allow_html=True)
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

    # 상단 헤더 및 다운로드 버튼
    header_col, action_col = st.columns([1, 0.4]) # action_col 넓이 확보
    with header_col:
        st.subheader("분석 결과")
    with action_col:
        dl_name = st.session_state.get("download_name", "result_오류검출.xlsx")
        data, out_name = build_excel_bytes(
            st.session_state.get("df_first_payment", pd.DataFrame()),
            st.session_state["df_errors"],
            dl_name,
            st.session_state.get("df_summary", pd.DataFrame()),
        )
        
        # 버튼 그룹 (다운로드 + 새로 분석하기)
        b_dl, b_new = st.columns([1.5, 1], gap="small")
        with b_dl:
            st.download_button(
                label="결과 엑셀 다운로드",
                data=data,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with b_new:
            if st.button("새로 분석하기", use_container_width=True):
                reset_app()
    
    # 요약 정보 렌더링
    st.success(f"✅ 처리가 완료되었습니다. (소요 시간: {duration:.2f}초)")
    
    # 노르딕 브루탈리스트 KPI 위젯
    st.markdown(f"""
        <div style="display: flex; gap: 24px; margin-bottom: 32px; padding: 0 40px;">
            <div class="nordic-card" style="flex: 1;">
                <p style="font-size: 12px; font-weight: 900; color: #9CA3AF; margin-bottom: 8px;">분석 대상</p>
                <p style="font-size: 32px; font-weight: 900; color: #1F2937; margin: 0;">{period_count:,}</p>
            </div>
            <div class="nordic-card-accent" style="flex: 1;">
                <p style="font-size: 12px; font-weight: 900; color: #1F2937; margin-bottom: 8px;">미납</p>
                <p style="font-size: 32px; font-weight: 900; color: #1F2937; margin: 0;">{c_miss} 건</p>
            </div>
            <div class="nordic-card" style="flex: 1;">
                <p style="font-size: 12px; font-weight: 900; color: #9CA3AF; margin-bottom: 8px;">부족</p>
                <p style="font-size: 32px; font-weight: 900; color: #1F2937; margin: 0;">{c_under} 건</p>
            </div>
            <div class="nordic-card" style="flex: 1;">
                <p style="font-size: 12px; font-weight: 900; color: #9CA3AF; margin-bottom: 8px;">초과</p>
                <p style="font-size: 32px; font-weight: 900; color: #1F2937; margin: 0;">{c_over} 건</p>
            </div>
            <div class="nordic-card-accent" style="flex: 1;">
                <p style="font-size: 12px; font-weight: 900; color: #1F2937; margin-bottom: 8px;">오류 합계</p>
                <p style="font-size: 32px; font-weight: 900; color: #1F2937; margin: 0;">{total_errors} 건</p>
            </div>
        </div>
    """, unsafe_allow_html=True)
    
    if c_filt > 0:
        st.info(f"💡 **참고:** {c_filt}건의 미납 내역이 '최초 납부월 이전'이라 제외되었습니다.")
    elif total_errors == 0:
        st.info("검출된 오류나 미납 내역이 없습니다.")


