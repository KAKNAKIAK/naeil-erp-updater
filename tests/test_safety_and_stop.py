import threading
import time
import tkinter as tk
import unittest
from unittest import mock

import excel_loader
from gui import RpaGuiApp, ensure_window_visible


class SafetyAndStopTest(unittest.TestCase):
    def test_ensure_window_visible_restores_iconic_window(self):
        window = mock.MagicMock()
        window.state.return_value = 'iconic'

        ensure_window_visible(window)

        window.deiconify.assert_called_once()
        window.lift.assert_called_once()
        window.attributes.assert_called_with('-topmost', True)
        window.after_idle.assert_called_once()

    def test_ensure_window_visible_handles_none_safely(self):
        ensure_window_visible(None)

    def test_select_excel_file_uses_parent_when_provided(self):
        parent_mock = mock.MagicMock()
        with mock.patch('excel_loader.filedialog.askopenfilename', return_value='C:/fake/path.xlsx') as mock_dialog:
            with mock.patch('excel_loader.tk.Tk') as mock_tk:
                result = excel_loader.select_excel_file(parent=parent_mock)
                self.assertEqual(result, 'C:/fake/path.xlsx')
                mock_dialog.assert_called_once_with(
                    parent=parent_mock,
                    title="요금 업데이트 엑셀 파일 선택",
                    filetypes=[("Excel Files", "*.xlsx *.xls")]
                )
                mock_tk.assert_not_called()

    def test_select_excel_file_creates_and_destroys_standalone_root_when_parent_none(self):
        with mock.patch('excel_loader.filedialog.askopenfilename', return_value='C:/fake/path.xlsx') as mock_dialog:
            with mock.patch('excel_loader.tk.Tk') as mock_tk:
                root_instance = mock.MagicMock()
                mock_tk.return_value = root_instance
                result = excel_loader.select_excel_file(parent=None)
                self.assertEqual(result, 'C:/fake/path.xlsx')
                mock_tk.assert_called_once()
                root_instance.withdraw.assert_called_once()
                root_instance.destroy.assert_called_once()

    def test_import_excel_to_sheet_ignores_duplicate_when_already_loading(self):
        app = mock.MagicMock()
        app._is_loading_excel = True
        app.is_running = False

        RpaGuiApp.import_excel_to_sheet(app)

        app.set_status.assert_not_called()

    def test_import_excel_to_sheet_locks_inputs_while_parsing(self):
        app = mock.MagicMock()
        app._is_loading_excel = False
        app.is_running = False
        app.root = mock.MagicMock()

        with mock.patch('gui.excel_loader.select_excel_file', return_value='C:/fake/path.xlsx'):
            with mock.patch('gui.threading.Thread') as thread_mock:
                RpaGuiApp.import_excel_to_sheet(app)

        self.assertTrue(app._is_loading_excel)
        app._set_inputs_locked.assert_called_once_with(True)
        app.start_btn.config.assert_called_once_with(state=tk.DISABLED)
        thread_mock.return_value.start.assert_called_once()

    def test_sleep_interruptible_exits_immediately_when_stopped(self):
        app = mock.MagicMock()
        app.is_running = False

        start = time.time()
        res = RpaGuiApp._sleep_interruptible(app, 5.0, check_interval=0.01)
        elapsed = time.time() - start

        self.assertFalse(res)
        self.assertLess(elapsed, 0.5)

    def test_sleep_interruptible_sleeps_full_duration_when_not_stopped(self):
        app = mock.MagicMock()
        app.is_running = True

        start = time.time()
        res = RpaGuiApp._sleep_interruptible(app, 0.05, check_interval=0.01)
        elapsed = time.time() - start

        self.assertTrue(res)
        self.assertGreaterEqual(elapsed, 0.04)

    def test_topas_wait_prompt_input_returns_none_on_stop_request(self):
        app = mock.MagicMock()
        app.topas_stop_requested = True
        app.is_running = True
        driver = mock.MagicMock()

        start = time.time()
        res = RpaGuiApp._topas_wait_prompt_input(app, driver, timeout=5.0)
        elapsed = time.time() - start

        self.assertIsNone(res)
        self.assertLess(elapsed, 0.5)

    def test_topas_wait_next_blocks_returns_none_on_stop_request_without_timeout_error(self):
        app = mock.MagicMock()
        app.topas_stop_requested = True
        app.is_running = True
        app._float_config.side_effect = lambda key, default, *args: default
        driver = mock.MagicMock()
        prev = mock.MagicMock()

        start = time.time()
        res = RpaGuiApp._topas_wait_next_blocks(app, driver, prev, expected_count=2, timeout=5.0)
        elapsed = time.time() - start

        self.assertIsNone(res)
        self.assertLess(elapsed, 0.5)

    def test_finish_topas_query_on_ui_handles_normal_stop_without_error_messagebox(self):
        app = mock.MagicMock()
        app.accent_orange = '#orange'
        app.accent_green = '#green'

        with mock.patch('gui.messagebox.showerror') as mock_err:
            RpaGuiApp.finish_topas_query_on_ui(app, results=['SAMPLE RAW'], error_message=None, stopped=True, elapsed=2.5)
            mock_err.assert_not_called()
            app.clean_up_ui_after_topas.assert_called_once()
            app._set_topas_status.assert_called_once_with('사용자 중지 · 부분 결과 보관', '#orange')
            app._apply_topas_results_to_slot.assert_called_once_with(['SAMPLE RAW'], stopped=True)

    def test_generate_and_show_report_restores_window_and_passes_parent(self):
        app = mock.MagicMock()
        app.root = mock.MagicMock()
        app.root.state.return_value = 'iconic'

        with mock.patch('gui.messagebox.showinfo') as mock_info:
            RpaGuiApp.generate_and_show_report(app, [{'date': '2026-06-15', 'status': 'SUCCESS'}])
            app.root.deiconify.assert_called_once()
            mock_info.assert_called_once_with(
                "요금 수정 완료",
                "전체 1일을 모두 정상적으로 수정했습니다.",
                parent=app.root,
            )

    def test_import_excel_to_sheet_skips_callback_when_closing(self):
        app = mock.MagicMock()
        app._is_loading_excel = False
        app.is_running = False
        app.is_closing = True
        app.root = mock.MagicMock()

        captured_worker = []
        def fake_thread_init(target=None, **kwargs):
            if target:
                captured_worker.append(target)
            t = mock.MagicMock()
            return t

        with mock.patch('gui.excel_loader.select_excel_file', return_value='C:/fake/path.xlsx'):
            with mock.patch('gui.excel_loader.load_fare_jobs_from_excel', return_value={'detected': False}):
                with mock.patch('gui.excel_loader.load_and_validate_fares', return_value=([], False)):
                    with mock.patch('gui.threading.Thread', side_effect=fake_thread_init):
                        RpaGuiApp.import_excel_to_sheet(app)
                        self.assertEqual(len(captured_worker), 1)
                        captured_worker[0]()
                        app.root.after.assert_not_called()

    def test_navigate_to_grid_page_returns_false_when_not_running(self):
        app = mock.MagicMock()
        app.is_running = False
        res = RpaGuiApp.navigate_to_grid_page(app, {}, 2, 5)
        self.assertFalse(res)

if __name__ == '__main__':
    unittest.main()
