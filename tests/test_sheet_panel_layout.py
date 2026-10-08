import ast
import inspect
import textwrap
import tkinter as tk
from tkinter import ttk
import unittest

from gui import RpaGuiApp


class SheetPanelLayoutTest(unittest.TestCase):
    def test_long_conditions_preserve_input_panel_width(self):
        root = tk.Tk()
        try:
            root.geometry('1006x800+10000+10000')
            app = RpaGuiApp.__new__(RpaGuiApp)
            app.root = root
            app.bg_color = '#ffffff'
            app.panel_width = 878
            app.collapsed_width = 1006
            app.expanded_width = 1884
            app.win_height = 800
            app.panel_expanded = False
            # Run the actual layout prefix, without initializing ERP or the rest of the app.
            method = ast.parse(textwrap.dedent(inspect.getsource(RpaGuiApp.build_ui))).body[0]
            prefix = []
            for statement in method.body:
                if isinstance(statement, ast.Assign) and any(
                    isinstance(target, ast.Name) and target.id == 'header_frame'
                    for target in statement.targets
                ):
                    break
                prefix.append(statement)
            exec(compile(ast.Module(body=prefix, type_ignores=[]), 'layout', 'exec'),
                 {'tk': tk}, {'self': app})
            app.toggle_btn = tk.Button(app.main_col)
            app.job_tree = ttk.Treeview(app.main_col,
                                       columns=('condition', 'status', 'rows', 'source'))
            app.job_tree.pack(fill=tk.X)
            app.job_queue = [{'hotel_name': 'Short hotel', 'rows': []}]
            app._refresh_job_queue_view()
            app.toggle_sheet_panel()
            root.update()
            original = (app.main_col.winfo_width(), app.side_panel.winfo_width())
            self.assertEqual(original[1], 878)

            app.job_queue[0]['hotel_name'] = 'Long hotel name ' * 40
            app._refresh_job_queue_view()
            tk.Label(app.main_col, text='Long status message ' * 100).pack()
            root.update()
            self.assertEqual((app.main_col.winfo_width(), app.side_panel.winfo_width()), original)

            app.toggle_sheet_panel()
            root.update()
            self.assertFalse(app.side_panel.winfo_ismapped())
            app.toggle_sheet_panel()
            root.update()
            self.assertEqual(app.side_panel.winfo_width(), 878)

            root.geometry('1400x800')
            root.update()
            self.assertEqual(app.side_panel.winfo_width(), 878)
            self.assertGreater(app.main_col.winfo_width(), 0)
        finally:
            root.destroy()
