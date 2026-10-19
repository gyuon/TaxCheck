from nicegui import app, run, ui
from nicegui.events import UploadEventArguments
from data_processor import (
    get_google_sheets_title,
    extract_first_payment_month,
    detect_errors,
    generate_missed_months,
    to_excel_bytes,
    normalize_names,
    extract_date_from_filename,
    filter_error_detection_target_payments,
    summarize_review_targets,
)
from constants import Col, Status
from excel_preparation import ExcelPreparation, ExcelPreparationStatus
from io import BytesIO
from datetime import datetime
from pathlib import Path
from typing import cast
import logging
import os
import pandas as pd
import time
from dotenv import load_dotenv

pd.set_option("future.no_silent_downcasting", True)

load_dotenv()
DEFAULT_GSHEET_URL = os.environ.get("GSHEET_URL", "")

PROJECT_DIR = Path(__file__).parent
app.add_static_files("/fonts", PROJECT_DIR / "assets" / "fonts")

RESULT_MAIN_CLASSES = (
    "result-main w-full max-w-[42rem] mx-auto my-auto p-5 md:p-6 gap-3 "
    "bg-white rounded-[2rem] shadow-sm hover:shadow-md transition-shadow"
)

APP_VIEW_CSS = """
@font-face {
    font-family: 'Pretendard Variable'; font-style: normal; font-weight: 100 900; font-display: swap;
    src: url('/fonts/PretendardVariable.woff2') format('woff2');
}
:root { --app-font-family: 'Pretendard Variable', Pretendard, sans-serif; }
body { background-color: #f9fafb; font-family: var(--app-font-family); }
button, input, textarea, select, .q-btn, .q-field, .q-tooltip, .q-notification {
    font-family: var(--app-font-family);
}
.app-shell { width: 100%; min-height: 100vh; padding: 24px; display: flex; }
.app-frame { width: 100%; max-width: 42rem; margin: auto; display: flex; flex-direction: column; gap: 12px; }
.app-version-footer { width: 100%; text-align: center; color: #887589; font-size: 13px;
    line-height: 18px; font-weight: 400; font-variant-numeric: tabular-nums; }
@media (max-width: 600px) { .app-shell { padding: 16px; } }
"""

INITIAL_VIEW_CSS = """
.initial-main { width: 100%; max-width: 42rem; margin: auto; padding: 28px;
    gap: 18px; align-items: stretch; background: white; border: 1px solid #dce1e8; border-radius: 8px; }
.initial-header { min-height: 52px; display: flex; flex-direction: column; align-items: center; text-align: center; }
.initial-title { font-size: 24px; line-height: 1.35; font-weight: 700; color: #18212d; word-break: keep-all; }
.initial-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 20px; }
.initial-upload-section { grid-column: 1; grid-row: 1 / span 2; min-width: 0; gap: 0; }
.initial-input-card { width: 100%; min-width: 0; padding: 20px; border: 1px solid #dce3ed;
    border-radius: 6px; background: #f8fafc; color: #18212d; font: inherit; text-align: left; }
.initial-upload-card { position: relative; height: 238px; cursor: pointer; border-style: dashed;
    display: flex; flex-direction: column; align-items: flex-start; justify-content: flex-start; }
.initial-card-title { font-size: 16px; line-height: 1.5; font-weight: 600; }
.initial-upload-body { position: absolute; inset: 0 16px; padding-top: 48px; display: grid;
    grid-template-rows: 40px 66px; align-content: center; justify-items: center; align-items: center; gap: 6px; }
.initial-upload-info { width: 100%; min-width: 0; display: flex; flex-direction: column;
    align-items: center; align-self: start; gap: 4px; }
.initial-upload-caption { max-width: 100%; text-align: center; font-size: 16px; line-height: 22px;
    color: #718096; overflow-wrap: anywhere; display: -webkit-box; -webkit-line-clamp: 2;
    -webkit-box-orient: vertical; overflow: hidden; }
.initial-upload-card.is-ready .initial-upload-caption { color: #18212d; font-weight: 600; }
.initial-upload-size { max-width: 100%; text-align: center; font-size: 14px; line-height: 18px; color: #8793a5; }
.initial-input-card.is-ready { background: #effbf3; border-color: #9edcb1; border-style: solid; }
.initial-status-icon { flex-shrink: 0; color: #4caf50; }
.initial-upload-icon { color: #94a3b8; }
.initial-period-card, .initial-sheet-card { height: 109px; min-height: 109px; padding: 16px 20px;
    display: flex; align-items: center; gap: 16px; }
.initial-period-card { cursor: pointer; }
.initial-period-icon { flex-shrink: 0; color: #94a3b8; }
.initial-card-copy { min-width: 0; gap: 4px; }
.initial-card-subtitle { max-width: 100%; color: #718096; font-size: 14px; line-height: 18px; overflow-wrap: anywhere; }
.initial-sheet-section { width: 100%; min-width: 0; gap: 0; }
.initial-sheet-card.is-failed { background: #fff5f5; border-color: #f0b4b4; }
.initial-sheet-card.is-failed .initial-status-icon { color: #d95151; }
.initial-sheet-card.is-failed .initial-card-subtitle { color: #b84444; }
.initial-sheet-card.is-loading .initial-status-icon { color: #8793a5; }
.initial-sheet-retry { color: #b84444; font-size: 13px; min-height: 24px; padding: 0 4px; }
.initial-upload-card:hover, .initial-period-card:hover { border-color: #8186e8; }
.initial-upload-card:focus-visible, .initial-period-card:focus-visible { outline: 2px solid #6368dc; outline-offset: 3px; }
.initial-analyze { width: 100%; min-height: 56px; border-radius: 6px; font-size: 16px;
    background-color: #3b41e3; color: white; }
.initial-analyze.q-btn.disabled { background-color: #b4bdca !important; color: white !important; opacity: 1 !important; }
.initial-period-dialog { width: 28rem; max-width: calc(100vw - 32px); padding: 24px; }
@media (max-width: 600px) {
    .initial-main { padding: 20px; gap: 20px; }
    .initial-title { font-size: 24px; }
    .initial-grid { grid-template-columns: minmax(0, 1fr); gap: 12px; }
    .initial-upload-section { grid-column: auto; grid-row: auto; }
    .initial-upload-card { height: 220px; }
    .initial-period-card, .initial-sheet-card { height: 96px; min-height: 96px; padding: 16px; }
}
"""


def _get_version() -> str:
    try:
        return Path(PROJECT_DIR, "VERSION").read_text().strip()
    except Exception:
        return "N/A"


def render_version_footer():
    with ui.element("footer").classes("app-version-footer"):
        ui.label(f"v{_get_version()}")


@ui.page("/")
async def index():
    ui.add_head_html(
        '<link rel="preload" href="/fonts/PretendardVariable.woff2" as="font" type="font/woff2" crossorigin>'
        '<style>'
        + APP_VIEW_CSS +
        ".nicegui-content { padding: 0 !important; }"
        ".q-btn.pulse-icon.disabled { opacity: 1 !important; cursor: wait; animation: button-pulse 1s ease-in-out infinite; }"
        "@keyframes button-pulse { 0%, 100% { background-color: #3b41e3; } 50% { background-color: #7b80f5; } }"
        + INITIAL_VIEW_CSS + "</style>"
    )

    is_result = app.storage.user.get("result_ready")

    with ui.element("div").classes("app-shell"):
        with ui.element("div").classes("app-frame"):
            main = ui.column().classes(RESULT_MAIN_CLASSES if is_result else "initial-main")
            render_version_footer()

    if is_result:
        with main:
            build_result_view_from_storage()
        return

    file_state = {"bytes": None, "name": None}
    settings = {
        "gsheet_url": DEFAULT_GSHEET_URL,
        "start_year": 2013,
        "end_year": datetime.now().year,
        "gsheet_ready": False,
    }
    analyze_button = None

    def update_analyze_button():
        if analyze_button is not None:
            analyze_button.set_enabled(bool(file_state["bytes"]) and settings["gsheet_ready"])

    with main:
        with ui.element("header").classes("initial-header"):
            ui.label("인별납부내역 오류검출").classes("initial-title")

        with ui.element("div").classes("initial-grid"):
            upload_section = ui.column().classes("initial-upload-section")
            with upload_section:
                render_upload_content(file_state, upload_section, update_analyze_button)
            render_period_card(settings)
            render_gsheet_card(settings, update_analyze_button)

        async def on_analyze():
            await analyze(file_state, settings, main)

        analyze_button = ui.button(
            "분석 실행",
            on_click=on_analyze,
            icon="play_arrow",
            color=None,
        ).props("unelevated no-caps").classes("initial-analyze")
        update_analyze_button()


def render_upload_content(file_state, parent, on_change=None):
    uploaded = bool(file_state["bytes"])
    card = ui.element("button").classes(
        "initial-input-card initial-upload-card" + (" is-ready" if uploaded else "")
    ).props(f'type=button aria-label="납부내역 파일 {"다시 업로드" if uploaded else "업로드"}"')
    with card:
        ui.label("납부내역 파일").classes("initial-card-title")
        with ui.element("div").classes("initial-upload-body"):
            ui.icon(
                "check_circle" if uploaded else "cloud_upload",
                size="36px",
            ).classes("initial-status-icon" if uploaded else "initial-upload-icon")
            with ui.element("div").classes("initial-upload-info"):
                caption = ui.label(
                    file_state["name"] if uploaded else "클릭하여 업로드 (.xlsx)"
                ).classes("initial-upload-caption")
                if uploaded:
                    caption.tooltip(file_state["name"])
                    size_kb = len(file_state["bytes"]) / 1024
                    size_text = f"{size_kb:,.1f} KB" if size_kb < 1024 else f"{size_kb / 1024:,.1f} MB"
                    ui.label(size_text).classes("initial-upload-size")

    uploader = ui.upload(
        auto_upload=True,
        on_upload=lambda e: on_file_uploaded(e, file_state, parent, on_change),
    ).props('accept=".xlsx"').classes("hidden")

    card.on("click", lambda: uploader.run_method("pickFiles"))


async def on_file_uploaded(e: UploadEventArguments, file_state, parent, on_change=None):
    file_state["bytes"] = await e.file.read()
    file_state["name"] = e.file.name
    if parent:
        parent.clear()
        with parent:
            render_upload_content(file_state, parent, on_change)
    if on_change is not None:
        on_change()


def render_gsheet_card(settings, on_change=None):
    container = ui.column().classes("initial-sheet-section")

    def render(title=None, loading=False):
        container.clear()
        configured = bool(settings["gsheet_url"])
        state = "is-loading" if loading else "is-ready" if title else "is-failed"
        with container:
            with ui.element("div").classes(f"initial-input-card initial-sheet-card {state}"):
                if loading:
                    ui.spinner(size="36px").classes("initial-status-icon")
                else:
                    ui.icon("check_circle" if title else "error", size="36px").classes("initial-status-icon")
                with ui.column().classes("initial-card-copy"):
                    ui.label("졸업생 명단").classes("initial-card-title")
                    subtitle = (
                        "명단을 불러오는 중" if loading else title if title else
                        "명단을 불러오지 못했습니다" if configured else "GSHEET_URL이 설정되지 않았습니다"
                    )
                    ui.label(subtitle).classes("initial-card-subtitle")
                    if configured and not title and not loading:
                        ui.button("다시 연결", icon="refresh", on_click=refresh, color=None).props(
                            "flat dense no-caps"
                        ).classes("initial-sheet-retry")

    async def refresh():
        settings["gsheet_ready"] = False
        render(loading=True)
        if on_change is not None:
            on_change()
        title = await run.io_bound(get_google_sheets_title, settings["gsheet_url"])
        settings["gsheet_ready"] = bool(title)
        render(title=title)
        if on_change is not None:
            on_change()

    settings["gsheet_ready"] = False
    render(loading=bool(settings["gsheet_url"]))
    if settings["gsheet_url"]:
        ui.timer(0.1, refresh, once=True)


def render_period_card(settings):
    cur_year = datetime.now().year
    with ui.element("button").props('type=button aria-label="분석 기간 설정"').classes(
        "initial-input-card initial-period-card"
    ) as card:
        ui.icon("event", size="36px").classes("initial-period-icon")
        with ui.column().classes("initial-card-copy"):
            ui.label("분석 기간").classes("initial-card-title")
            period_subtitle = ui.label(
                f"{settings['start_year']}년 ~ {settings['end_year']}년"
            ).classes("initial-card-subtitle")

    async def open_dialog():
        with ui.dialog().classes("rounded-xl") as dialog:
            with ui.card().classes("initial-period-dialog"):
                ui.label("분석 기간 설정").classes("text-lg font-semibold mb-4")
                with ui.row().classes("w-full gap-4"):
                    start_input = ui.number(
                        "시작 년도",
                        value=settings["start_year"],
                        min=2000,
                        max=cur_year,
                        step=1,
                    ).classes("flex-1 min-w-0")
                    end_input = ui.number(
                        "종료 년도",
                        value=settings["end_year"],
                        min=2000,
                        max=cur_year + 10,
                        step=1,
                    ).classes("flex-1 min-w-0")
                with ui.row().classes("w-full justify-end gap-2 mt-4"):
                    ui.button("취소", on_click=dialog.close).props("outline")
                    ui.button("저장", on_click=lambda: save()).props("unelevated")

        def save():
            if start_input.value is None or end_input.value is None:
                ui.notify("시작 년도와 종료 년도를 입력하세요.", type="warning")
                return
            if start_input.value > end_input.value:
                ui.notify("시작 년도는 종료 년도보다 클 수 없습니다.", type="warning")
                return
            settings["start_year"] = int(start_input.value)
            settings["end_year"] = int(end_input.value)
            period_subtitle.set_text(
                f"{settings['start_year']}년 ~ {settings['end_year']}년"
            )
            dialog.close()

        dialog.open()

    card.on("click", open_dialog)


async def analyze(file_state, settings, main):
    if not file_state["bytes"] or not settings["gsheet_url"]:
        ui.notify("원본 엑셀과 Google Sheets URL을 입력하세요.", type="warning")
        return

    main.classes(remove="initial-main", add=RESULT_MAIN_CLASSES)
    main.clear()
    with main:
        with ui.column().classes("w-full p-12 flex flex-col items-center"):
            ui.spinner("dots").classes("mb-4")
            ui.label("데이터 처리 중입니다").classes("text-slate-500")

    try:
        df_errors, df_first, df_summary, dl_name, summary, df_view = await run.io_bound(
            process_data,
            file_state["bytes"],
            file_state["name"],
            settings["gsheet_url"],
            settings["start_year"],
            settings["end_year"],
        )
    except Exception as e:
        import traceback
        print(f"[ERROR] 분석 실패: {e}\n{traceback.format_exc()}")

        try:
            main.clear()
            with main:
                ui.label(f"오류 발생: {e}").classes("text-red-500")
                with ui.expansion("기술적 상세 정보"):
                    ui.label(traceback.format_exc()).classes(
                        "text-xs text-slate-400 whitespace-pre-wrap"
                    )
        except RuntimeError:
            print("[ERROR] RuntimeError: UI 업데이트 실패 (클라이언트 연결 끊김 가능성)")
        return

    serialize_start = time.time()
    df_errors_records = df_errors.to_dict("records")
    df_first_records = df_first.to_dict("records")
    df_summary_records = df_summary.to_dict("records") if not df_summary.empty else []
    df_view_records = df_view.to_dict("records")
    logging_log = logging.getLogger("taxcheck")
    logging_log.info(
        "결과 직렬화 완료 (%.2f초), df_errors=%d건, df_first=%d건, df_summary=%d건, df_view=%d건",
        time.time() - serialize_start,
        len(df_errors_records),
        len(df_first_records),
        len(df_summary_records),
        len(df_view_records),
    )

    app.storage.user["result"] = {
        "df_errors": df_errors_records,
        "df_first": df_first_records,
        "df_summary": df_summary_records,
        "df_view": df_view_records,
        "dl_name": dl_name,
        "summary": summary,
    }
    app.storage.user["result_ready"] = True

    try:
        main.clear()
        with main:
            build_result_view_from_storage()
    except RuntimeError:
        print("[ERROR] RuntimeError: 결과 뷰 UI 업데이트 실패")


def process_data(file_bytes, file_name, gsheet_url, start_year, end_year):
    import logging
    log = logging.getLogger("taxcheck")
    log.setLevel(logging.DEBUG)
    if not log.handlers:
        h = logging.StreamHandler()
        h.setFormatter(logging.Formatter("[%(levelname)s] %(message)s"))
        log.addHandler(h)

    start_time = time.time()
    log.info("1/8 엑셀 파일 읽기: %s", file_name)

    try:
        excel_read_start = time.time()
        df = pd.read_excel(
            BytesIO(file_bytes),
            sheet_name="raw",
            header=1,
            engine="calamine",
            usecols=["이름", "해당년", "해당월", "코드1", "코드2", "입금"],
        )
    except Exception as e:
        raise ValueError(f"원본 엑셀 읽기 오류: {e}")
    log.info(
        "1/8 엑셀 파일 읽기 완료 (%.2f초), rows=%d, cols=%d",
        time.time() - excel_read_start,
        len(df),
        len(df.columns),
    )

    column_prep_start = time.time()
    rename_map = {"해당년": Col.YEAR, "해당월": Col.MONTH}
    df = df.rename(columns=rename_map)

    required_cols = [Col.NAME, Col.YEAR, Col.MONTH, Col.CODE1, Col.CODE2, Col.RAW_DEPOSIT]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"누락된 열: {missing}")

    df = normalize_names(df)
    log.info(
        "1/8 컬럼 정리 완료 (%.2f초), required_cols=%d",
        time.time() - column_prep_start,
        len(required_cols),
    )

    numeric_start = time.time()
    numeric_columns = [Col.YEAR, Col.MONTH, Col.CODE1, Col.CODE2, Col.RAW_DEPOSIT]
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].round().astype("Int64")
    log.info(
        "1/8 숫자 컬럼 변환 완료 (%.2f초), numeric_cols=%d",
        time.time() - numeric_start,
        len(numeric_columns),
    )

    def to_csv_url(url: str) -> str:
        if "export?format=csv" in url:
            return url
        if "/edit" in url:
            return url.split("/edit")[0] + "/export?format=csv"
        return url

    csv_url = to_csv_url(gsheet_url)
    log.info("2/8 Google Sheets CSV 다운로드: %s", csv_url)
    csv_start = time.time()
    df_sheet = pd.read_csv(csv_url)
    log.info(
        "2/8 Google Sheets CSV 다운로드 완료 (%.2f초), rows=%d, cols=%d",
        time.time() - csv_start,
        len(df_sheet),
        len(df_sheet.columns),
    )
    df_sheet = normalize_names(df_sheet)

    if "구분" not in df_sheet.columns or "이름" not in df_sheet.columns:
        raise ValueError("Google Sheet에 '이름'과 '구분' 열이 필요합니다.")

    graduation_names = list(
        set(df_sheet[df_sheet["구분"].str.strip() == "졸업생"]["이름"].tolist())
    )

    exemption_map: dict[str, list[str]] = {}
    if "면제기금" in df_sheet.columns:
        for _, row in df_sheet[df_sheet["구분"].str.strip() == "졸업생"].iterrows():
            name = str(row["이름"]).strip()
            exempt_str = str(row.get("면제기금", "")).strip()
            if name and exempt_str:
                exempt_funds = [f.strip() for f in exempt_str.split(",") if f.strip()]
                if exempt_funds:
                    exemption_map[name] = exempt_funds

    log.info("3/8 졸업생 명단 필터링: %d명", len(graduation_names))
    df = df[df["이름"].isin(graduation_names)].copy()
    total_grad_count = len(df)

    if df.empty:
        raise ValueError("졸업생 명단 필터링 후 남은 데이터가 없습니다.")

    before_target_filter_count = len(df)
    df = filter_error_detection_target_payments(cast(pd.DataFrame, df))
    log.info(
        "   오류검출 대상 코드 필터링: %d건 → %d건",
        before_target_filter_count,
        len(df),
    )

    if df.empty:
        raise ValueError("오류검출 대상 코드 필터링 후 남은 데이터가 없습니다.")

    df_first_payment = extract_first_payment_month(cast(pd.DataFrame, df))
    log.info("4/8 최초납부월 추출 완료")

    if Col.YEAR in df.columns:
        mask = (df[Col.YEAR] >= start_year) & (df[Col.YEAR] <= end_year)
        mask = mask.fillna(False)
        df_view: pd.DataFrame = cast(pd.DataFrame, df[mask].copy())
        period_count = len(df_view)
        log.info("5/8 년도 필터링: %d ~ %d → %d건", start_year, end_year, period_count)
        if df_view.empty:
            raise ValueError(
                f"선택한 년도 범위({start_year} ~ {end_year})에 해당하는 데이터가 없습니다."
            )
    else:
        df_view = cast(pd.DataFrame, df.copy())
        period_count = len(df_view)

    log.info("6/8 오류 검출 시작")
    detect_start = time.time()
    df_errors = detect_errors(df_view)
    log.info(
        "6/8 오류 검출 완료 (%.2f초), df_view=%d건, df_errors=%d건",
        time.time() - detect_start,
        len(df_view),
        len(df_errors),
    )

    ref_date = extract_date_from_filename(file_name)
    ref_year = ref_date[0] if ref_date else end_year
    ref_month = ref_date[1] if ref_date else 12

    effective_end_year = min(ref_year, end_year)
    effective_end_month = ref_month if ref_year <= end_year else 12

    log.info("7/8 미납월 생성")
    missed_start = time.time()
    df_missed, filtered_count = generate_missed_months(
        df_view,
        df_first_payment,
        filename=file_name,
        graduation_names=graduation_names,
        start_year=start_year,
        start_month=1,
        end_year=effective_end_year,
        end_month=effective_end_month,
        exemption_map=exemption_map,
    )
    log.info(
        "7/8 미납월 생성 완료 (%.2f초), df_missed=%d건, filtered_count=%d, graduates=%d",
        time.time() - missed_start,
        len(df_missed),
        filtered_count,
        len(graduation_names),
    )

    if not df_missed.empty:
        if df_errors.empty:
            df_errors = df_missed
        else:
            for col in [Col.DEPOSIT, Col.STANDARD, Col.DIFF]:
                if col in df_errors.columns:
                    df_errors[col] = cast(
                        pd.Series, pd.to_numeric(df_errors[col], errors="coerce")
                    ).astype(float)
                if col in df_missed.columns:
                    df_missed[col] = cast(
                        pd.Series, pd.to_numeric(df_missed[col], errors="coerce")
                    ).astype(float)
            for col in [Col.CODE]:
                if col in df_errors.columns:
                    df_errors[col] = df_errors[col].astype(str)
                if col in df_missed.columns:
                    df_missed[col] = df_missed[col].astype(str)

            df_missed_aligned = df_missed.reindex(columns=df_errors.columns)
            df_errors = pd.concat([df_errors, df_missed_aligned], ignore_index=True)

    if not df_errors.empty:
        df_errors = df_errors.sort_values(
            [Col.NAME, Col.YEAR, Col.MONTH, Col.CODE]
        ).reset_index(drop=True)
        df_errors.index = df_errors.index + 1

    download_name = file_name.replace(".xlsx", "_오류검출.xlsx")

    df_summary = pd.DataFrame()
    if not df_errors.empty:
        summary_by_name = cast(
            pd.DataFrame,
            df_errors.groupby([Col.NAME, Col.STATUS])
            .size()
            .unstack(fill_value=0)
            .astype(int)
            .reset_index(),
        )
        for status in [Status.UNPAID, Status.INSUFFICIENT, Status.EXCESS]:
            if status not in summary_by_name.columns:
                summary_by_name[status] = 0
        summary_by_name = cast(
            pd.DataFrame,
            summary_by_name[
                [Col.NAME, Status.UNPAID, Status.INSUFFICIENT, Status.EXCESS]
            ],
        )
        summary_by_name["합계"] = (
            summary_by_name[Status.UNPAID]
            + summary_by_name[Status.INSUFFICIENT]
            + summary_by_name[Status.EXCESS]
        )
        summary_by_name = cast(
            pd.DataFrame, summary_by_name[summary_by_name["합계"] > 0]
        )
        summary_by_name = cast(
            pd.DataFrame,
            summary_by_name.sort_values(
                ["합계", Status.UNPAID, Status.INSUFFICIENT, Status.EXCESS, Col.NAME],
                ascending=[False, False, False, False, True],
            ),
        ).reset_index(drop=True)
        summary_by_name.index = summary_by_name.index + 1
        df_summary = summary_by_name.rename(
            columns={
                Col.NAME: "이름",
                Status.UNPAID: "미납",
                Status.INSUFFICIENT: "부족",
                Status.EXCESS: "초과",
            }
        )

    counts = {Status.UNPAID: 0, Status.INSUFFICIENT: 0, Status.EXCESS: 0}
    if not df_errors.empty:
        type_counts = df_errors[Col.STATUS].value_counts().to_dict()
        for k in counts.keys():
            counts[k] = int(type_counts.get(k, 0))

    summary_info = {
        "counts": counts,
        "filtered_count": filtered_count,
        "total_grad_count": total_grad_count,
        "period_count": period_count,
        "duration": time.time() - start_time,
        "period": f"{start_year}년 ~ {end_year}년",
        "start_year": int(start_year),
        "end_year": int(end_year),
    }
    log.info("8/8 완료 (%.2f초), df_errors=%d건, df_summary=%d건",
             summary_info["duration"], len(df_errors), len(df_summary))

    return df_errors, df_first_payment, df_summary, download_name, summary_info, df_view


def build_result_view_from_storage():
    r = app.storage.user["result"]
    df_errors = pd.DataFrame(r["df_errors"])
    df_first = pd.DataFrame(r["df_first"])
    df_summary = pd.DataFrame(r["df_summary"]) if r["df_summary"] else pd.DataFrame()
    df_view = pd.DataFrame(r.get("df_view", []))
    build_result_view(df_errors, df_first, df_summary, r["dl_name"], r["summary"], df_view)


def build_result_view(df_errors, df_first, df_summary, dl_name, summary, df_view=None):
    review = summarize_review_targets(df_errors)
    duration = summary["duration"]
    period_count = summary["period_count"]
    c_filt = summary["filtered_count"]

    ui.label("분석 결과").classes("text-2xl font-bold text-slate-800 text-center w-full")

    with ui.element("div").classes("result-summary w-full border border-slate-200"):
        with ui.element("dl").classes("grid grid-cols-2 divide-x divide-slate-200 m-0"):
            for label, value, unit in [
                ("오류 대상", review["target_count"], "건"),
                ("대상자", review["person_count"], "명"),
            ]:
                with ui.element("div").classes("min-w-0 text-center"):
                    with ui.element("dt").classes(
                        "bg-slate-100 border-b border-slate-200 px-3 py-1.5 text-xs sm:text-sm font-medium text-slate-600"
                    ):
                        ui.label(label)
                    with ui.element("dd").classes(
                        "m-0 px-3 py-2.5 text-2xl sm:text-3xl font-semibold text-slate-900 tabular-nums"
                    ):
                        ui.label(f"{value:,}{unit}")

        with ui.element("table").classes(
            "w-full table-fixed border-collapse border-t border-slate-200 text-slate-800"
        ).props('aria-label="오류 대상 상태별 집계"'):
            with ui.element("thead").classes("bg-slate-100 text-slate-600"):
                with ui.element("tr"):
                    for heading in ["구분", "오류 대상", "3개월 이상 미납"]:
                        alignment = "text-left" if heading == "구분" else "text-center"
                        with ui.element("th").props('scope="col"').classes(
                            f"{alignment} border-b border-slate-200 px-3 py-2 text-xs sm:text-sm font-medium leading-snug"
                        ):
                            label = ui.label(heading)
                            if heading == "3개월 이상 미납":
                                label.tooltip("기준월보다 3개월 이상 이전의 미납 건수")
            with ui.element("tbody"):
                for status in [Status.UNPAID, Status.INSUFFICIENT, Status.EXCESS]:
                    border = "border-b border-slate-200" if status != Status.EXCESS else ""
                    with ui.element("tr").classes(border):
                        with ui.element("th").props('scope="row"').classes(
                            "px-3 py-1.5 text-left text-sm font-medium"
                        ):
                            ui.label(status)
                        with ui.element("td").classes(
                            "px-3 py-1.5 text-center text-base font-semibold tabular-nums"
                        ):
                            ui.label(f"{review['counts'][status]:,}건")
                        overdue = f"{review['long_unpaid_count']:,}건" if status == Status.UNPAID else "-"
                        foreground = "text-orange-600" if status == Status.UNPAID else "text-slate-400"
                        with ui.element("td").classes(
                            f"{foreground} px-3 py-1.5 text-center text-base font-semibold tabular-nums"
                        ):
                            ui.label(overdue)

    with ui.element("dl").classes("result-metadata self-stretch mx-px my-0 border-t border-slate-200"):
        for label, value in [
            ("분석 기간", summary.get("period", "-")),
            ("분석한 납부내역", f"{period_count:,}건"),
            ("처리 시간", f"{duration:.2f}초"),
        ]:
            with ui.element("div").classes(
                "grid grid-cols-3 border-b border-slate-200 py-1.5"
            ):
                with ui.element("dt").classes("col-span-2 px-3 text-xs sm:text-sm text-slate-600"):
                    ui.label(label)
                with ui.element("dd").classes("min-w-0 m-0 px-2 text-center text-sm text-slate-700 tabular-nums break-keep"):
                    ui.label(value)

    if review["target_count"] == 0:
        with ui.row().classes("items-center gap-2 text-green-700"):
            ui.icon("check_circle", size="20px")
            ui.label("확인할 오류가 없습니다.").classes("text-sm")

    if c_filt > 0:
        ui.label(f"{c_filt:,}건의 미납 내역이 최초 납부월 이전이라 제외되었습니다.").classes(
            "text-xs text-gray-500"
        )

    with ui.column().classes("result-actions w-full gap-2"):
        preparation = ExcelPreparation()
        initial_generation = preparation.start()
        initial_presentation = preparation.presentation
        download_button = ui.button(
            initial_presentation.text,
            icon=initial_presentation.icon,
            color=None,
        ).props("size=md unelevated").style("background-color: #3b41e3; color: white; opacity: 1 !important; font-size: 14px; border-radius: 8px;").classes("w-full min-h-[40px]")

        def update_download_button():
            presentation = preparation.presentation
            download_button.set_text(presentation.text)
            download_button.props(f"icon={presentation.icon}")
            download_button.props(f"aria-busy={'true' if presentation.pulsing else 'false'}")
            if presentation.pulsing:
                download_button.classes(add="pulse-icon")
            else:
                download_button.classes(remove="pulse-icon")
            if presentation.disabled:
                download_button.disable()
            else:
                download_button.enable()

        update_download_button()

        async def prepare_excel(generation: int):
            generation_started = time.perf_counter()
            result_log = logging.getLogger("taxcheck")
            try:
                excel_data, _ = await run.io_bound(
                    to_excel_bytes,
                    df_first,
                    df_errors,
                    dl_name,
                    df_summary,
                    df_view,
                    summary.get("start_year"),
                    summary.get("end_year"),
                )
            except Exception as e:
                result_log.exception(
                    "결과 엑셀 생성 실패 (%.2f초)",
                    time.perf_counter() - generation_started,
                )
                if preparation.fail(generation, str(e)):
                    try:
                        update_download_button()
                        ui.notify(f"엑셀 생성 실패: {e}", type="negative")
                    except RuntimeError:
                        result_log.info("엑셀 생성 실패 후 클라이언트 연결 종료")
                return

            if not preparation.complete(generation, excel_data):
                result_log.info("이전 엑셀 생성 결과 폐기")
                return

            result_log.info(
                "결과 엑셀 생성 완료 (%.2f초), bytes=%d",
                time.perf_counter() - generation_started,
                len(excel_data),
            )
            try:
                update_download_button()
            except RuntimeError:
                result_log.info("엑셀 생성 완료 후 클라이언트 연결 종료")

        async def handle_download():
            if preparation.status == ExcelPreparationStatus.READY and preparation.data is not None:
                ui.download(
                    preparation.data,
                    dl_name,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
                return

            if preparation.status == ExcelPreparationStatus.FAILED:
                retry_generation = preparation.start()
                update_download_button()
                await prepare_excel(retry_generation)

        download_button.on("click", handle_download)
        ui.timer(0.15, lambda: prepare_excel(initial_generation), once=True)

        async def reset():
            preparation.invalidate()
            app.storage.user["result_ready"] = False
            ui.navigate.to("/")

        ui.button(
            "새로 분석하기",
            on_click=reset,
            icon="refresh",
        ).props("outline size=md color=grey-7").style("border-radius: 8px; font-size: 14px;").classes("w-full min-h-[40px]")




@app.get("/health", status_code=200)
async def health():
    return {"status": "ok"}

ui.run(
    title="인별납부내역 오류검출",
    host="0.0.0.0",
    port=int(os.environ.get("PORT", 10010)),
    reload=os.environ.get("RENDER") != "true",
    storage_secret=os.environ.get("STORAGE_SECRET", "taxcheck-local-dev"),
)
