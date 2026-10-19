import importlib
import unittest
from contextlib import contextmanager
from types import SimpleNamespace
from unittest.mock import AsyncMock, Mock, patch

from nicegui import ui
from nicegui.client import Client
from test_sheet_connection import response_with

with patch.object(ui, "run"):
    view = importlib.import_module("app")


def labels(parent):
    return [node.text for node in parent.descendants() if isinstance(node, ui.label)]


class InitialViewTestCase(unittest.IsolatedAsyncioTestCase):
    @contextmanager
    def ui_parent(self):
        client = Client(ui.page("/_initial_view_test"))
        self.addCleanup(client.delete)
        with client:
            with ui.column() as parent:
                yield parent


class TestInitialUploadView(InitialViewTestCase):
    def test_empty_upload_displays_prompt_and_card_title(self):
        with self.ui_parent() as parent:
            view.render_upload_content({"bytes": None, "name": None}, parent)

        self.assertEqual(labels(parent), ["납부내역 파일", "클릭하여 업로드 (.xlsx)"])
        card = parent.default_slot.children[0]
        self.assertEqual(card.tag, "button")
        self.assertEqual(card._props["type"], "button")
        self.assertNotIn("is-ready", card._classes)

    def test_uploaded_file_displays_name_and_size_without_extra_action(self):
        filename = "정리본_26.8.19.xlsx"
        with self.ui_parent() as parent:
            view.render_upload_content({"bytes": b"xlsx", "name": filename}, parent)

        self.assertEqual(labels(parent), ["납부내역 파일", filename, "0.0 KB"])
        self.assertIn("is-ready", parent.default_slot.children[0]._classes)
        icon = next(node for node in parent.descendants() if isinstance(node, ui.icon))
        self.assertEqual(icon._props["name"], "check_circle")
        self.assertEqual(icon._props["size"], "36px")
        self.assertEqual(parent.default_slot.children[0]._props["aria-label"], "납부내역 파일 다시 업로드")

    def test_upload_size_uses_kilobytes_and_megabytes(self):
        for byte_count, size_text in [(4096, "4.0 KB"), (1024 * 1024, "1.0 MB"), (3 * 1024 * 1024, "3.0 MB")]:
            with self.subTest(byte_count=byte_count):
                with self.ui_parent() as parent:
                    view.render_upload_content({"bytes": b"x" * byte_count, "name": "sample.xlsx"}, parent)
                self.assertEqual(labels(parent), ["납부내역 파일", "sample.xlsx", size_text])

    async def test_reupload_replaces_old_file_and_updates_button_state(self):
        state = {"bytes": b"old", "name": "old.xlsx"}
        changed = Mock()
        with self.ui_parent() as parent:
            view.render_upload_content(state, parent, changed)

        event = SimpleNamespace(file=SimpleNamespace(name="new.xlsx", read=AsyncMock(return_value=b"new")))
        await view.on_file_uploaded(event, state, parent, changed)

        self.assertEqual(state, {"bytes": b"new", "name": "new.xlsx"})
        self.assertEqual(labels(parent), ["납부내역 파일", "new.xlsx", "0.0 KB"])
        self.assertEqual(len(parent.default_slot.children), 2)
        changed.assert_called_once_with()


class TestInitialSheetView(InitialViewTestCase):
    async def test_direct_csv_connection_enables_readiness(self):
        settings = {"gsheet_url": "https://example.invalid/export?format=csv"}
        with patch.object(ui, "timer") as timer:
            with self.ui_parent() as parent:
                view.render_gsheet_card(settings)

        response = response_with("이름,구분\n홍길동,졸업생\n")
        with patch("data_processor.requests.get", return_value=response):
            with patch.object(view.run, "io_bound", new=AsyncMock(side_effect=lambda func, *args: func(*args))):
                await timer.call_args.args[1]()

        self.assertTrue(settings["gsheet_ready"])
        self.assertEqual(labels(parent), ["졸업생 명단", "졸업생 명단"])

    async def test_connection_loading_then_success(self):
        settings = {"gsheet_url": "https://example.invalid/sheet"}
        changed = Mock()
        with patch.object(ui, "timer") as timer:
            with self.ui_parent() as parent:
                view.render_gsheet_card(settings, changed)

        self.assertFalse(settings["gsheet_ready"])
        self.assertIn("명단을 불러오는 중", labels(parent))
        with patch.object(view.run, "io_bound", new=AsyncMock(return_value="졸업생 명단 원본")):
            await timer.call_args.args[1]()

        self.assertTrue(settings["gsheet_ready"])
        self.assertEqual(labels(parent), ["졸업생 명단", "졸업생 명단 원본"])
        icon = next(node for node in parent.descendants() if isinstance(node, ui.icon))
        self.assertEqual(icon._props["name"], "check_circle")
        self.assertEqual(icon._props["size"], "36px")
        self.assertEqual(changed.call_count, 2)

    async def test_failed_connection_can_be_retried(self):
        settings = {"gsheet_url": "https://example.invalid/sheet"}
        with patch.object(ui, "timer") as timer:
            with self.ui_parent() as parent:
                view.render_gsheet_card(settings)

        refresh = timer.call_args.args[1]
        with patch.object(view.run, "io_bound", new=AsyncMock(side_effect=[None, "재연결된 명단"])):
            await refresh()
            self.assertFalse(settings["gsheet_ready"])
            self.assertIn("명단을 불러오지 못했습니다", labels(parent))
            retry = next(node for node in parent.descendants() if isinstance(node, ui.button))
            self.assertEqual(retry.text, "다시 연결")
            self.assertTrue(any(listener.type == "click" for listener in retry._event_listeners.values()))
            await refresh()

        self.assertTrue(settings["gsheet_ready"])
        self.assertEqual(labels(parent), ["졸업생 명단", "재연결된 명단"])
        self.assertFalse(any(isinstance(node, ui.button) for node in parent.descendants()))

    def test_missing_url_does_not_attempt_connection(self):
        settings = {"gsheet_url": "", "gsheet_ready": True}
        with patch.object(ui, "timer") as timer:
            with self.ui_parent() as parent:
                view.render_gsheet_card(settings)

        self.assertFalse(settings["gsheet_ready"])
        self.assertIn("GSHEET_URL이 설정되지 않았습니다", labels(parent))
        timer.assert_not_called()


class TestInitialPeriodView(InitialViewTestCase):
    def test_period_icon_matches_upload_check_icon_size(self):
        with self.ui_parent() as parent:
            view.render_period_card({"start_year": 2013, "end_year": 2026})

        icon = next(node for node in parent.descendants() if isinstance(node, ui.icon))
        self.assertEqual(icon._props["name"], "event")
        self.assertEqual(icon._props["size"], "36px")


class TestSharedPageView(InitialViewTestCase):
    def test_version_footer_contains_only_version_text(self):
        with self.ui_parent() as parent:
            view.render_version_footer()

        footer = parent.default_slot.children[0]
        self.assertEqual(footer.tag, "footer")
        self.assertIn("app-version-footer", footer._classes)
        self.assertEqual(labels(parent), [f"v{view._get_version()}"])
        self.assertFalse(any(isinstance(node, ui.icon) for node in parent.descendants()))

    def test_footer_survives_main_content_replacement(self):
        with self.ui_parent() as parent:
            with ui.column() as main:
                ui.label("초기 화면")
            view.render_version_footer()

        for screen in ["데이터 처리 중입니다", "분석 결과", "오류 발생"]:
            main.clear()
            with main:
                ui.label(screen)
            self.assertEqual(labels(parent), [screen, f"v{view._get_version()}"])

    def test_local_font_asset_and_license_are_bundled(self):
        fonts = view.PROJECT_DIR / "assets" / "fonts"
        font = fonts / "PretendardVariable.woff2"
        self.assertEqual(font.read_bytes()[:4], b"wOF2")
        self.assertGreater(font.stat().st_size, 1_000_000)
        self.assertIn("SIL OPEN FONT LICENSE", (fonts / "Pretendard-LICENSE.txt").read_text())

    async def test_analyze_keeps_column_layout_and_footer_on_error(self):
        with self.ui_parent() as parent:
            main = ui.column().classes("initial-main")
            view.render_version_footer()

        with patch.object(view.run, "io_bound", new=AsyncMock(side_effect=ValueError("invalid file"))):
            with patch("builtins.print"):
                await view.analyze(
                    {"bytes": b"invalid", "name": "invalid.xlsx"},
                    {"gsheet_url": "https://example.invalid/sheet", "start_year": 2013, "end_year": 2026},
                    main,
                )

        self.assertIn("nicegui-column", main._classes)
        self.assertNotIn("initial-main", main._classes)
        self.assertIn("오류 발생: invalid file", labels(parent))
        self.assertEqual(labels(parent).count(f"v{view._get_version()}"), 1)

    def test_common_style_uses_local_variable_font(self):
        self.assertIn("@font-face", view.APP_VIEW_CSS)
        self.assertIn("url('/fonts/PretendardVariable.woff2')", view.APP_VIEW_CSS)
        self.assertIn("font-weight: 100 900", view.APP_VIEW_CSS)
        self.assertIn("font-family: var(--app-font-family)", view.APP_VIEW_CSS)
        self.assertIn("text-align: center", view.APP_VIEW_CSS)


if __name__ == "__main__":
    unittest.main()
