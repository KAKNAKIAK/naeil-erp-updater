import io
import sys
import tkinter as tk
import unittest
from unittest import mock

from tksheet import Sheet

from gui import RpaGuiApp, SHEET_HEADERS


class PendingSheetInputTest(unittest.TestCase):
    def setUp(self):
        self.root = tk.Tk()
        self.root.withdraw()

    def tearDown(self):
        self.root.destroy()

    def _make_app(self):
        app = RpaGuiApp.__new__(RpaGuiApp)
        app.root = self.root
        app.sheet = Sheet(
            self.root,
            data=[['20261001', '', '', '', '', '', '', '']],
            headers=SHEET_HEADERS,
        )
        app.sheet.pack()
        self.root.update_idletasks()
        app.is_running = False
        app.job_queue = []
        app.period_mode_var = tk.BooleanVar(self.root, value=False)
        app.filter_mode = tk.StringVar(self.root, value='ALL')
        app.fb_var = tk.StringVar(self.root, value='')
        app.formula_entry = tk.Entry(self.root, textvariable=app.fb_var)
        app._active_cell = (0, 1)
        app.formulas = {}
        app._results = {}
        app._recalc_busy = False
        app._applying_sheet_snapshot = False
        app._sheet_last_snapshot = None
        app._ref_mode = False
        app._snapshot_sheet_state = mock.Mock(return_value={})
        app._record_sheet_undo_state = mock.Mock()
        app._sync_sheet_undo_baseline = mock.Mock()
        app._try_import_erp_conditions_for_direct_run = mock.Mock(return_value={})
        app._set_inputs_locked = mock.Mock()
        app.set_status = mock.Mock()
        for name in ('start_btn', 'pause_btn', 'stop_btn', 'progress_bar',
                     'progress_lbl', 'log_txt', 'active_lbl'):
            setattr(app, name, mock.MagicMock())
        app.accent_orange = 'orange'
        app.accent_green = 'green'
        app.fg_muted = 'gray'
        app.sheet.extra_bindings([('modified', app._on_sheet_modified)])
        return app

    def _invoke_start(self, app, focused_widget):
        original_stdout = sys.stdout
        button = tk.Button(self.root, command=app.start_rpa)
        try:
            # A withdrawn test window has no OS focus. Preserve the widget that
            # would retain focus when a Tk button is pressed in the real app.
            with mock.patch.object(self.root, 'focus_get', return_value=focused_widget), \
                 mock.patch.object(app.sheet.MT, 'focus_get', return_value=focused_widget), \
                 mock.patch('gui.threading.Thread'), \
                 mock.patch('gui.GUIConsoleRedirector', return_value=io.StringIO()):
                button.invoke()
        finally:
            sys.stdout = original_stdout
            button.destroy()

    def _assert_requested_change(self, app, text, expected_amount):
        row = app.fares_data[0]
        self.assertEqual(row['date'], '2026-10-01')
        self.assertEqual(row['adult_air'], expected_amount)
        if text == '예약마감':
            self.assertEqual(row['progress_status'], '예약마감')
            self.assertEqual(row['progress_status_field'], 'adult_air')
            self.assertEqual(app._progress_status_from_text(row['progress_status']), ('05', '예약마감'))
        else:
            self.assertEqual(row['progress_status'], '')

    def test_start_commits_active_cell_before_reading_status_amount_or_formula(self):
        for text, expected_amount in (('예약마감', ''), ('420000', 420000), ('=1+2', 3)):
            with self.subTest(text=text):
                app = self._make_app()
                app.sheet.MT.open_text_editor(r=0, c=1)
                app.sheet.MT.text_editor.window.set_text(text)
                self.assertEqual(app.sheet.get_sheet_data()[0][1], '')

                self._invoke_start(app, app.sheet.MT.text_editor.tktext)

                self.assertFalse(app.sheet.MT.text_editor.open)
                self._assert_requested_change(app, text, expected_amount)
                app.sheet.destroy()
                app.formula_entry.destroy()

    def test_start_commits_focused_formula_bar_before_reading_status_amount_or_formula(self):
        for text, expected_amount in (('예약마감', ''), ('420000', 420000), ('=1+2', 3)):
            with self.subTest(text=text):
                app = self._make_app()
                app.fb_var.set(text)
                self.assertEqual(app.sheet.get_sheet_data()[0][1], '')

                self._invoke_start(app, app.formula_entry)

                self._assert_requested_change(app, text, expected_amount)
                app.sheet.destroy()
                app.formula_entry.destroy()


if __name__ == '__main__':
    unittest.main()
