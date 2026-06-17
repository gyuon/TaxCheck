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
)
from constants import Col, Status
from io import BytesIO
from datetime import datetime
from pathlib import Path
from typing import cast
import os
import pandas as pd
import time
from dotenv import load_dotenv

pd.set_option("future.no_silent_downcasting", True)

load_dotenv()
DEFAULT_GSHEET_URL = os.environ.get("GSHEET_URL", "")

PROJECT_DIR = Path(__file__).parent


def _get_version() -> str:
    try:
        return Path(PROJECT_DIR, "VERSION").read_text().strip()
    except Exception:
        return "N/A"


@ui.page("/")
async def index():
    ui.add_head_html(
        '<style>'
        "body { background-color: #f9fafb; font-family: 'Inter', 'Pretendard', sans-serif; }"
        ".nicegui-content { padding: 0 !important; }"
        "</style>"
    )

    is_result = app.storage.user.get("result_ready")

    with ui.element("div").classes("w-full min-h-[calc(100vh-2rem)] flex flex-col items-center"):
        main = ui.column().classes(
            "w-full max-w-[42rem] mx-auto my-auto p-5 md:p-9 "
            "bg-white rounded-[2rem] shadow-sm hover:shadow-md transition-shadow"
        )

    if is_result:
        with main:
            build_result_view_from_storage()
        return

    file_state = {"bytes": None, "name": None}
    settings = {
        "gsheet_url": DEFAULT_GSHEET_URL,
        "start_year": 2013,
        "end_year": datetime.now().year,
    }

    with main:
        ui.label("인별납부내역 오류검출").classes(
            "text-3xl font-bold text-center text-slate-800 mb-0 w-full"
        )
        ui.label(f"v{_get_version()}").classes(
            "text-sm text-slate-400 text-center w-full"
        ).style("margin-bottom: 0.2rem")

        upload_section = ui.column().classes("w-full mb-2")
        with upload_section:
            render_upload_content(file_state, upload_section)

        render_gsheet_card(settings)
        render_period_card(settings)

        async def on_analyze():
            await analyze(file_state, settings, main)

        ui.button(
            "분석 실행",
            on_click=on_analyze,
            icon="play_arrow",
        ).props("size=lg unelevated").style(
            "background-color: #3b41e3 !important; color: white; font-size: 15px; border-radius: 8px; padding-top: 1rem; padding-bottom: 1rem;"
        ).classes("w-full mt-3")


def render_upload_content(file_state, parent):
    if file_state["bytes"]:
        size_kb = len(file_state["bytes"]) / 1024
        size_str = f"{size_kb:,.1f} KB" if size_kb < 1024 else f"{size_kb / 1024:,.1f} MB"
        card = ui.element("div").classes(
            "w-full border-2 border-solid border-green-200 bg-green-50 "
            "rounded-xl p-4 h-[142.7px] flex flex-col items-center justify-center gap-2 "
            "cursor-pointer hover:border-blue-400 hover:bg-green-100/60 transition-all"
        )
        with card:
            ui.icon("check_circle", size="28px", color="green")
            with ui.column().classes("items-center gap-0"):
                ui.label(file_state["name"]).classes("text-slate-800 font-semibold text-[15px]")
                ui.label(size_str).classes("text-slate-400 text-sm")
            ui.label("다시 업로드").classes("text-slate-500 text-sm underline")
    else:
        card = ui.element("div").classes(
            "w-full border-2 border-dashed border-slate-200 bg-slate-50/50 rounded-xl "
            "h-[142.7px] flex flex-col items-center justify-center gap-2 cursor-pointer "
            "hover:border-blue-400 hover:bg-blue-50 transition-all"
        )
        with card:
            ui.icon("cloud_upload", size="28px").classes("text-slate-400")
            ui.label("클릭하여 업로드 (.xlsx)").classes("text-slate-500 text-[15px]")

    uploader = ui.upload(
        auto_upload=True,
        on_upload=lambda e: on_file_uploaded(e, file_state, parent),
    ).props('accept=".xlsx"').classes("hidden")

    card.on("click", lambda: uploader.run_method("pickFiles"))


async def on_file_uploaded(e: UploadEventArguments, file_state, parent):
    file_state["bytes"] = await e.file.read()
    file_state["name"] = e.file.name
    if parent:
        parent.clear()
        with parent:
            render_upload_content(file_state, parent)


def render_gsheet_card(settings):
    container = ui.column().classes("w-full mb-2")

    def _render():
        container.clear()
        with container:
            url = settings["gsheet_url"]

            if not url:
                ui.element("div").classes(
                    "w-full border-2 border-dashed border-slate-200 bg-slate-50/50 rounded-xl "
                    "h-[142.7px] flex flex-col items-center justify-center gap-2"
                )
                with ui.element("div"):
                    ui.icon("table_chart", size="28px").classes("text-slate-400")
                    ui.label("GSHEET_URL이 설정되지 않았습니다").classes(
                        "text-slate-400 text-[15px]"
                    )
                return

            title = get_google_sheets_title(url)

            if title:
                with ui.element("div").classes(
                    "w-full border-2 border-solid border-green-200 bg-green-50 "
                    "rounded-xl p-4 h-[142.7px] flex flex-col items-center justify-center gap-2 "
                    "transition-all"
                ):
                    ui.icon("check_circle", size="28px", color="green")
                    with ui.column().classes("items-center gap-0"):
                        ui.label("졸업생 명단").classes("text-slate-800 font-semibold text-[15px]")
                        ui.label(title).classes("text-slate-400 text-sm")
            else:
                with ui.element("div").classes(
                    "w-full border-2 border-solid border-red-200 bg-red-50 "
                    "rounded-xl p-4 h-[142.7px] flex flex-col items-center justify-center gap-2 "
                    "cursor-pointer hover:border-red-400 hover:bg-red-100/60 transition-all"
                ) as card:
                    ui.icon("error", size="28px", color="red")
                    with ui.column().classes("items-center gap-0"):
                        ui.label("졸업생 명단").classes("text-slate-800 font-semibold text-[15px]")
                        ui.label("연결 실패. 다시 시도하려면 클릭하세요").classes(
                            "text-red-600 text-sm text-center"
                        )
                card.on("click", lambda: _render())

    _render()


def render_period_card(settings):
    cur_year = datetime.now().year
    with ui.card().classes(
        "w-full p-4 flex flex-col items-center cursor-pointer hover:scale-[1.01] transition-transform"
    ) as card:
        with ui.element("div").classes("bg-slate-50 p-2 rounded-lg shrink-0 mb-0.5"):
            ui.icon("event", size="24px").classes("text-slate-600")
        with ui.column().classes("gap-0 items-center"):
            ui.label("분석 기간").classes("text-slate-800 font-semibold text-base text-center")
            period_subtitle = ui.label(
                f"{settings['start_year']}년부터 {settings['end_year']}년까지"
            ).classes("text-slate-500 text-[15px] text-center")

    async def open_dialog():
        with ui.dialog().classes("rounded-xl") as dialog:
            with ui.card().classes("min-w-[28rem] p-6"):
                ui.label("분석 기간 설정").classes("text-lg font-semibold mb-4")
                with ui.row().classes("w-full gap-4"):
                    start_input = ui.number(
                        "시작 년도",
                        value=settings["start_year"],
                        min=2000,
                        max=cur_year,
                        step=1,
                    )
                    end_input = ui.number(
                        "종료 년도",
                        value=settings["end_year"],
                        min=2000,
                        max=cur_year + 10,
                        step=1,
                    )
                with ui.row().classes("w-full justify-end gap-2 mt-4"):
                    ui.button("취소", on_click=dialog.close).props("outline")
                    ui.button("저장", on_click=lambda: save()).props("unelevated")

        def save():
            settings["start_year"] = int(start_input.value)
            settings["end_year"] = int(end_input.value)
            period_subtitle.set_text(
                f"{settings['start_year']}년부터 {settings['end_year']}년까지"
            )
            dialog.close()

        dialog.open()

    card.on("click", open_dialog)


async def analyze(file_state, settings, main):
    if not file_state["bytes"] or not settings["gsheet_url"]:
        ui.notify("원본 엑셀과 Google Sheets URL을 입력하세요.", type="warning")
        return

    main.clear()
    with main:
        with ui.column().classes("w-full p-12 flex flex-col items-center"):
            ui.spinner("dots").classes("mb-4")
            ui.label("데이터 처리 중입니다...").classes("text-slate-500")

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
                        "text-xs text-slate-400 font-mono whitespace-pre-wrap"
                    )
        except RuntimeError:
            print("[ERROR] RuntimeError: UI 업데이트 실패 (클라이언트 연결 끊김 가능성)")
        return

    app.storage.user["result"] = {
        "df_errors": df_errors.to_dict("records"),
        "df_first": df_first.to_dict("records"),
        "df_summary": df_summary.to_dict("records") if not df_summary.empty else [],
        "df_view": df_view.to_dict("records"),
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
        df = pd.read_excel(BytesIO(file_bytes), sheet_name="raw", header=1)
    except Exception as e:
        raise ValueError(f"원본 엑셀 읽기 오류: {e}")

    rename_map = {"해당년": Col.YEAR, "해당월": Col.MONTH}
    df = df.rename(columns=rename_map)

    required_cols = [Col.NAME, Col.YEAR, Col.MONTH, Col.CODE1, Col.CODE2, Col.RAW_DEPOSIT]
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        raise ValueError(f"누락된 열: {missing}")

    df = normalize_names(df)

    numeric_columns = [Col.YEAR, Col.MONTH, Col.CODE1, Col.CODE2, Col.RAW_DEPOSIT]
    for col in numeric_columns:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")
            df[col] = df[col].round().astype("Int64")

    def to_csv_url(url: str) -> str:
        if "export?format=csv" in url:
            return url
        if "/edit" in url:
            return url.split("/edit")[0] + "/export?format=csv"
        return url

    csv_url = to_csv_url(gsheet_url)
    log.info("2/8 Google Sheets CSV 다운로드: %s", csv_url)
    df_sheet = pd.read_csv(csv_url)
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
    df_errors = detect_errors(df_view)
    log.info("   오류 검출 완료: %d건", len(df_errors))

    ref_date = extract_date_from_filename(file_name)
    ref_year = ref_date[0] if ref_date else end_year
    ref_month = ref_date[1] if ref_date else 12

    effective_end_year = min(ref_year, end_year)
    effective_end_month = ref_month if ref_year <= end_year else 12

    log.info("7/8 미납월 생성")
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
    counts = summary["counts"]
    c_miss = counts.get(Status.UNPAID, 0)
    c_under = counts.get(Status.INSUFFICIENT, 0)
    c_over = counts.get(Status.EXCESS, 0)
    total_errors = sum(counts.values())
    duration = summary["duration"]
    period_count = summary["period_count"]
    c_filt = summary["filtered_count"]

    ui.label("분석 결과").classes("text-2xl font-bold text-slate-800 text-center w-full mb-4")

    with ui.column().classes("w-full items-center gap-3 mb-8"):
        excel_data, _ = to_excel_bytes(df_first, df_errors, dl_name, df_summary, df_view)
        download_button = ui.button(
            "결과 엑셀 다운로드",
            icon="download",
        ).props("size=md unelevated").style("background-color: #3b41e3 !important; color: white; font-size: 14px; border-radius: 8px;").classes("w-full max-w-xs")
        download_button.on("click", lambda: ui.download(
            excel_data,
            dl_name,
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        ))
        async def reset():
            app.storage.user["result_ready"] = False
            ui.navigate.to("/")

        ui.button(
            "새로 분석하기",
            on_click=reset,
            icon="refresh",
        ).props("outline size=md color=grey-7").style("border-radius: 8px; font-size: 14px;").classes("w-full max-w-xs")

    kpi_configs = [
        ("분석 대상", period_count, "bg-slate-100", "bg-slate-300", "text-slate-500", "text-slate-800"),
        ("미납", c_miss, "bg-red-100", "bg-red-500", "text-red-800", "text-red-800"),
        ("부족", c_under, "bg-indigo-100", "bg-indigo-500", "text-indigo-800", "text-indigo-800"),
        ("초과", c_over, "bg-pink-100", "bg-pink-500", "text-pink-800", "text-pink-800"),
        ("오류 합계", total_errors, "bg-violet-100", "bg-[#232936]", "text-violet-700", "text-violet-800"),
    ]

    with ui.row().classes("w-full gap-3 mb-8"):
        for label, value, bg, bar, label_color, value_color in kpi_configs:
            with ui.element("div").classes(
                f"flex-1 rounded-xl py-5 px-3 text-center relative overflow-hidden shadow-sm {bg}"
            ):
                ui.element("div").classes(f"absolute top-0 left-0 right-0 h-1 {bar}")
                ui.label(label).classes(f"text-sm font-semibold {label_color} tracking-wide mb-2")
                ui.label(f"{value:,}").classes(f"text-3xl font-medium {value_color} tracking-tight")

    messages = [f"처리 시간: {duration:.2f}초"]
    if c_filt > 0:
        messages.append(f"{c_filt}건의 미납 내역이 '최초 납부월 이전'이라 제외되었습니다.")
    elif total_errors == 0:
        messages.append("검출된 오류나 미납 내역이 없습니다.")

    with ui.row().classes("w-full items-center gap-2 text-sm text-slate-600"):
        ui.icon("info", size="sm").classes("text-slate-400")
        ui.label("\n".join(messages)).classes("leading-relaxed")




@app.get("/health", status_code=200)
async def health():
    return {"status": "ok"}

ui.run(
    title="인별납부내역 오류검출",
    host="0.0.0.0",
    port=int(os.environ.get("PORT", 10000)),
    reload=True,
    storage_secret=os.environ.get("STORAGE_SECRET", "taxcheck-local-dev"),
)
