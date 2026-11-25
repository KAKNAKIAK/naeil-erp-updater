import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from unittest import mock

import openpyxl

import erp_cli
import gui


def _write_xlsx(path, rows):
    wb = openpyxl.Workbook()
    ws = wb.active
    for row in rows:
        ws.append(row)
    wb.save(path)
    wb.close()


class ErpCliTest(unittest.TestCase):
    def setUp(self):
        self.tmpdir = tempfile.TemporaryDirectory()

    def tearDown(self):
        self.tmpdir.cleanup()

    def _xlsx(self, name, rows):
        path = os.path.join(self.tmpdir.name, name)
        _write_xlsx(path, rows)
        return path

    def test_single_sheet_preserves_blank_columns_skips_errors_and_merges_with_gaps(self):
        path = self._xlsx("sabre_like.xlsx", [
            ["날짜", "운임"],
            ["2026-12-01", 210000],
            ["2026-12-02", 210000],
            ["2026-12-03", "조회불가"],
            ["2026-12-04", 210000],
            ["2026-12-05", "예약마감"],
            ["2026-12-06", "예약마감"],
        ])
        loaded = erp_cli.load_jobs_for_cli(
            path,
            price_desc="다낭",
            airline_code="lj",
            merge=True,
        )
        rows = loaded["fares_data"]
        self.assertEqual(len(rows), 3)
        self.assertEqual((rows[0]["date"], rows[0]["date_end"], rows[0]["adult_air"]), ("2026-12-01", "2026-12-02", "210000"))
        self.assertEqual(rows[0]["adult_hotel"], "")
        self.assertEqual(rows[0]["adult_profit"], "")
        self.assertEqual(rows[0]["airline_code"], "LJ")
        self.assertEqual((rows[1]["date"], rows[1]["date_end"], rows[1]["adult_air"]), ("2026-12-04", "2026-12-04", "210000"))
        self.assertEqual((rows[2]["date"], rows[2]["date_end"], rows[2]["adult_air"], rows[2]["progress_status"]),
                         ("2026-12-05", "2026-12-06", "", "예약마감"))
        self.assertTrue(any("조회불가" in w for w in loaded["warnings"]))

    def test_single_sheet_blocks_unfiltered_by_default(self):
        path = self._xlsx("unfiltered.xlsx", [
            ["날짜", "항공비"],
            ["2026-12-01", 150000],
        ])
        with self.assertRaisesRegex(ValueError, "최소 1개 조건"):
            erp_cli.load_jobs_for_cli(path)

    def test_multi_job_excel_plan_outputs_clean_json_to_stdout(self):
        path = self._xlsx("multi_jobs.xlsx", [
            ["요금구분", "항공사코드", "시작일", "종료일", "항공비"],
            ["다낭", "LJ", "2026-12-01", "2026-12-02", 200000],
            ["푸꾸옥", "VJ", "2026-12-05", "2026-12-05", "예약마감"],
        ])
        buf = io.StringIO()
        with redirect_stdout(buf):
            code = erp_cli.main(["--excel", path, "--plan", "--start", "2026-12-02"])
        self.assertEqual(code, 0)
        data = json.loads(buf.getvalue())
        self.assertTrue(data["ok"])
        self.assertEqual(data["mode"], "plan")
        self.assertEqual(data["jobsCount"], 2)
        self.assertEqual(data["planned"], 2)
        self.assertEqual(data["items"][0]["fares"], {"adult_air": "200000"})
        self.assertEqual(data["items"][1]["progress"], "예약마감")

    def test_explicit_erp_address_does_not_fallback_to_9222(self):
        runner = erp_cli.HeadlessErpRunner(
            {"debugger_address": "127.0.0.1:9222", "erp_debugger_address": "127.0.0.1:9223"},
            dry_run=True,
            explicit_erp_address="127.0.0.1:9224",
        )
        configs = runner._debug_browser_configs("ERP")
        self.assertEqual([c["address"] for c in configs], ["127.0.0.1:9224"])
        self.assertTrue(runner.config["dry_run"])

    def test_dry_run_never_calls_parent_save_or_progress_apply(self):
        cfg = erp_cli.load_cli_config()
        cfg["dry_run"] = False
        runner = erp_cli.HeadlessErpRunner(cfg, dry_run=True)
        self.assertTrue(runner.config["dry_run"])

        fake_driver = mock.MagicMock()
        runner._connect_matching_debug_browser = mock.MagicMock(
            return_value=(fake_driver, {"address": "127.0.0.1:9224"})
        )
        runner.find_and_switch_frame = mock.MagicMock(return_value=True)
        runner._set_erp_date_input = mock.MagicMock(side_effect=lambda _s, v, _l, **_k: v)
        runner._set_erp_airline_filter = mock.MagicMock(return_value={"text": "[LJ] 진에어"})
        runner._set_erp_text_filter = mock.MagicMock(return_value={"value": "다낭"})
        runner._set_erp_hotel_filter = mock.MagicMock(return_value={"value": "", "seq": ""})
        runner._query_erp_with_retry = mock.MagicMock(return_value=True)
        runner.get_grid_page_state = mock.MagicMock(return_value=(1, 1, 2, False))
        runner._current_page_progress_rows = mock.MagicMock(return_value=[
            {"basePriceSeq": "101", "procCd": "04"},
            {"basePriceSeq": "102", "procCd": "04"},
        ])
        runner._current_page_progress_status_counts = mock.MagicMock(
            return_value={"total": 2, "reservation_closed": 0}
        )
        runner._current_page_progress_status_matches = mock.MagicMock(return_value=False)

        rows = [
            {
                "date": "2026-12-01", "date_end": "2026-12-02",
                "adult_air": "210000", "adult_hotel": "", "adult_land": "",
                "adult_tour": "", "adult_profit": "", "child_fare": "", "infant_fare": "",
                "price_desc": "다낭", "airline_code": "LJ", "departure_flight": "",
                "hotel_name": "", "progress_status": "",
            },
            {
                "date": "2026-12-05", "date_end": "2026-12-06",
                "adult_air": "", "adult_hotel": "", "adult_land": "",
                "adult_tour": "", "adult_profit": "", "child_fare": "", "infant_fare": "",
                "price_desc": "다낭", "airline_code": "LJ", "departure_flight": "",
                "hotel_name": "", "progress_status": "예약마감",
            },
        ]

        with mock.patch.object(gui.RpaGuiApp, "_save_current_page") as parent_save, \
             mock.patch.object(gui.RpaGuiApp, "_apply_current_page_progress_status") as parent_prog:
            history = runner.run_fares(rows)

        parent_save.assert_not_called()
        parent_prog.assert_not_called()
        self.assertEqual(len(history), 2)
        self.assertTrue(all(h["status"] == "SUCCESS" for h in history))
        self.assertEqual([a["action"] for a in runner.dry_run_actions], ["price_update", "progress_status"])

    def test_startup_connection_error_returns_nonzero_json_without_popup(self):
        path = self._xlsx("one.xlsx", [
            ["요금구분", "항공사코드", "시작일", "종료일", "항공비"],
            ["다낭", "LJ", "2026-12-01", "2026-12-01", 200000],
        ])
        buf = io.StringIO()
        with mock.patch.object(
            erp_cli.HeadlessErpRunner,
            "_connect_matching_debug_browser",
            side_effect=RuntimeError("9224 연결 실패"),
        ), redirect_stdout(buf):
            code = erp_cli.main(["--excel", path, "--erp-address", "127.0.0.1:9224"])
        self.assertEqual(code, 1)
        data = json.loads(buf.getvalue())
        self.assertFalse(data["ok"])
        self.assertIn("9224 연결 실패", data["error"])

    def test_date_range_normalizes_compact_dates_and_rejects_invalid_or_reversed(self):
        rows = [{"date": "2026-12-05", "date_end": "2026-12-05"}]
        self.assertEqual(len(erp_cli.filter_rows_by_date_range(rows, start="20261201", end="2026/12/10")), 1)
        with self.assertRaisesRegex(ValueError, "--start"):
            erp_cli.filter_rows_by_date_range(rows, start="2026-02-30")
        with self.assertRaisesRegex(ValueError, "늦습니다"):
            erp_cli.filter_rows_by_date_range(rows, start="2026-12-10", end="2026-12-01")

    def test_auto_stop_in_paging_loop_preserves_auto_stop_error_not_user_stop(self):
        cfg = erp_cli.load_cli_config()
        runner = erp_cli.HeadlessErpRunner(cfg, dry_run=True)
        runner._connect_matching_debug_browser = mock.MagicMock(
            return_value=(mock.MagicMock(), {"address": "127.0.0.1:9224"})
        )
        runner.find_and_switch_frame = mock.MagicMock(return_value=True)
        runner._set_erp_date_input = mock.MagicMock(side_effect=lambda _s, v, _l, **_k: v)
        runner._set_erp_airline_filter = mock.MagicMock(return_value={"text": "[LJ] 진에어"})
        runner._set_erp_text_filter = mock.MagicMock(return_value={"value": "다낭"})
        runner._set_erp_hotel_filter = mock.MagicMock(return_value={"value": "", "seq": ""})
        runner._query_erp_with_retry = mock.MagicMock(return_value=True)
        runner.get_grid_page_state = mock.MagicMock(return_value=(1, 2, 600, True))
        runner._current_page_progress_status_counts = mock.MagicMock(
            return_value={"total": 500, "reservation_closed": 0}
        )
        runner._dismiss_alert_and_modal = mock.MagicMock()

        def fail_and_auto_stop(*_args, **_kwargs):
            runner.is_running = False
            raise RuntimeError("cleanup abort failed")

        runner.navigate_to_grid_page = mock.MagicMock(side_effect=fail_and_auto_stop)
        rows = [{
            "date": "2026-12-01", "date_end": "2026-12-01",
            "adult_air": "210000", "adult_hotel": "", "adult_land": "",
            "adult_tour": "", "adult_profit": "", "child_fare": "", "infant_fare": "",
            "price_desc": "다낭", "airline_code": "LJ", "departure_flight": "",
            "hotel_name": "", "progress_status": "",
        }]
        history = runner.run_fares(rows)
        self.assertEqual(len(history), 1)
        self.assertEqual(history[0]["status"], "FAIL")
        self.assertIn("자동 안전 중단", history[0]["error"])
        self.assertIn("cleanup abort failed", history[0]["error"])
        self.assertNotIn("사용자 중단", history[0]["error"])


if __name__ == "__main__":
    unittest.main()
