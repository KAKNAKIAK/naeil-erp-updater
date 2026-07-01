from datetime import date
from pathlib import Path
import tempfile
import unittest
from unittest import mock

from fare.calculator import RoundTripResult, calculate_round_trips, js_round
from fare.exporter import export_results_to_excel, results_to_tsv, to_erp_rows
from fare.parser import parse_topas_text
from gui import AIRLINE_EMPTY_LABEL, RpaGuiApp


DEP_TEXT = """>
AN14JUNICNGUM/ALJ915 -AC-
** AMADEUS AVAILABILITY - AN ** GUM GUAM.GU  3 SU 14JUN 0000
 1   LJ 915  Y9 W9 L9 P3 /ICN 2 GUM I 1815 2345 E0/738 4:30
>
AC1
AN15JUNICNGUM/ALJ915 -AC-
** AMADEUS AVAILABILITY - AN ** GUM GUAM.GU  4 MO 15JUN 0000
 1   LJ 915  Y9 W9 L5 /ICN 2 GUM I 1815 2345 E0/738 4:30
>
"""

RET_TEXT = """>
AN17JUNGUMICN/ALJ916 -AC-
** AMADEUS AVAILABILITY - AN ** SEL SEOUL.KR  6 WE 17JUN 0000
 1   LJ 916  Y9 W9 L9 /GUM I ICN 2 0135 0520 E0/738 4:45
>
AC1
AN18JUNGUMICN/ALJ916 -AC-
** AMADEUS AVAILABILITY - AN ** SEL SEOUL.KR  7 TH 18JUN 0000
 1   LJ 916  Y9 W4 /GUM I ICN 2 0135 0520 E0/738 4:45
>
"""

FARES = [
    {"route": "ICN-GUM-LJ", "type": "기준", "classCode": "Y", "roundTripFare": 1000000},
    {"route": "ICN-GUM-LJ", "type": "기준", "classCode": "W", "roundTripFare": 600000},
    {"route": "ICN-GUM-LJ", "type": "기준", "classCode": "L", "roundTripFare": 400000},
    {"route": "ICN-GUM-LJ", "type": "LOW", "classCode": "Y", "roundTripFare": 900000},
]

SEASONS = [
    {"route": "ICN-GUM-LJ", "type": "LOW", "startDate": "2026-06-15", "endDate": "2026-06-20"}
]


class FareEngineTest(unittest.TestCase):
    def test_parser_uses_future_year_and_available_4_to_9_only(self):
        parsed = parse_topas_text(DEP_TEXT, today=date(2026, 6, 11))

        self.assertEqual([row.date for row in parsed.records], ["2026-06-14", "2026-06-15"])
        self.assertEqual(parsed.records[0].classes, ("Y9", "W9", "L9"))
        self.assertEqual(parsed.records[1].classes, ("Y9", "W9", "L5"))

    def test_parser_duplicate_date_last_wins(self):
        raw = "AN10JUNICNGUM/ALJ915\nY9\nAN10JUNICNGUM/ALJ915\nW9"
        parsed = parse_topas_text(raw, today=date(2026, 6, 11))

        self.assertEqual(parsed.records[0].date, "2027-06-10")
        self.assertEqual(parsed.records[0].classes, ("W9",))
        self.assertTrue(any("중복 날짜" in warning for warning in parsed.warnings))

    def test_calculates_by_nights_and_marks_closed_without_exporting_amount(self):
        result = calculate_round_trips(
            DEP_TEXT,
            RET_TEXT,
            FARES,
            SEASONS,
            "ICN-GUM-LJ",
            nights=(3,),
            today=date(2026, 6, 11),
        )

        rows = result.result[3]
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0].dep_date, "2026-06-14")
        self.assertEqual(rows[0].ret_date, "2026-06-17")
        self.assertFalse(rows[0].is_closed)
        self.assertEqual(rows[0].total_fare, 200000 + 200000)
        self.assertEqual(results_to_tsv(rows).splitlines()[0], "2026-06-14\t400000")

    def test_closed_rows_are_excluded_from_erp_rows(self):
        closed_dep = "AN14JUNICNGUM/ALJ915\nP3"
        result = calculate_round_trips(
            closed_dep,
            RET_TEXT,
            FARES,
            SEASONS,
            "ICN-GUM-LJ",
            nights=(3,),
            today=date(2026, 6, 11),
        )

        rows = result.result[3]
        self.assertTrue(rows[0].is_closed)
        self.assertEqual(results_to_tsv(rows), "2026-06-14\t마감")
        self.assertEqual(to_erp_rows(rows, profit=50000), [])

    def test_erp_rows_fill_only_air_fare_by_default(self):
        result = calculate_round_trips(
            DEP_TEXT,
            RET_TEXT,
            FARES,
            SEASONS,
            "ICN-GUM-LJ",
            nights=(3,),
            today=date(2026, 6, 11),
        )

        rows = to_erp_rows(result.result[3])

        self.assertEqual(rows[0]["adult_air"], 400000)
        self.assertEqual(rows[0]["adult_hotel"], "")
        self.assertEqual(rows[0]["adult_land"], "")
        self.assertEqual(rows[0]["adult_tour"], "")
        self.assertEqual(rows[0]["adult_profit"], "")
        self.assertEqual(rows[0]["child_fare"], "")
        self.assertEqual(rows[0]["infant_fare"], "")

    def test_sheet_reservation_closed_text_becomes_progress_status(self):
        class Var:
            def __init__(self, value):
                self.value = value

            def get(self):
                return self.value

        class Sheet:
            def get_sheet_data(self):
                return [["2026-07-15", "예약마감", "", "", "", "", "", ""]]

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.period_mode_var = Var(False)
        app.sheet = Sheet()

        rows, errors = app.read_sheet_data()

        self.assertEqual(errors, [])
        self.assertEqual(rows[0]["progress_status"], "예약마감")
        self.assertEqual(rows[0]["progress_status_field"], "adult_air")
        self.assertEqual(rows[0]["adult_air"], "")
        self.assertEqual(app._progress_status_from_text(rows[0]["progress_status"]), ("05", "예약마감"))

    def test_load_job_to_sheet_restores_reservation_closed_cell_text(self):
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class Sheet:
            def __init__(self):
                self.data = []

            def headers(self, _headers):
                pass

            def set_column_widths(self, _widths):
                pass

            def set_sheet_data(self, data, **_kwargs):
                self.data = data

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.period_mode_var = Var(False)
        app.price_desc_var = Var()
        app.hotel_name_var = Var()
        app.progress_text_var = Var()
        app.sheet = Sheet()
        app.formulas = {}
        app._results = {}
        app.panel_expanded = True
        app._record_sheet_undo_state = lambda *_args, **_kwargs: None
        app._clear_merge_restore_snapshot = lambda: None
        app._select_airline_code = lambda *_args, **_kwargs: None
        app._set_source_badge = lambda *_args, **_kwargs: None
        app.refresh_count = lambda: None
        app._load_active_into_fb = lambda: None
        app._sync_sheet_undo_baseline = lambda: None

        app._load_job_to_sheet({
            "price_desc": "괌_저녁_3박",
            "airline_code": "LJ",
            "hotel_name": "두짓비치 괌",
            "rows": [{
                "date": "2026-07-15",
                "date_end": "2026-07-15",
                "adult_air": "",
                "progress_status": "예약마감",
                "progress_status_field": "adult_air",
            }],
        })

        self.assertEqual(app.sheet.data[0][1], "예약마감")
        self.assertEqual(app.hotel_name_var.get(), "두짓비치 괌")

    def test_current_page_progress_counts_detects_reservation_closed(self):
        class Driver:
            def execute_script(self, *_args):
                return [
                    {"procCd": "05", "procNm": "예약마감"},
                    {"procCd": "04", "procNm": "예약신청"},
                    {"procCd": "", "procNm": "예약 마감"},
                ]

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.config = {"grid_id": "#gridMain"}
        app.driver = Driver()

        summary = app._current_page_progress_status_counts()

        self.assertEqual(summary["total"], 3)
        self.assertEqual(summary["reservation_closed"], 2)
        self.assertEqual(summary["counts"]["05|예약마감"], 1)
        self.assertFalse(RpaGuiApp._should_skip_price_update_for_all_closed(summary))

    def test_price_update_skips_only_when_all_rows_are_reservation_closed(self):
        self.assertFalse(
            RpaGuiApp._should_skip_price_update_for_all_closed(
                {"total": 3, "reservation_closed": 2}
            )
        )
        self.assertTrue(
            RpaGuiApp._should_skip_price_update_for_all_closed(
                {"total": 3, "reservation_closed": 3}
            )
        )
        self.assertFalse(
            RpaGuiApp._should_skip_price_update_for_all_closed(
                {"total": 0, "reservation_closed": 0}
            )
        )

    def test_job_normalizes_korean_airline_label_to_code(self):
        app = RpaGuiApp.__new__(RpaGuiApp)
        app._set_airline_choices([("LJ", "진에어"), ("7C", "제주항공")])

        job = app._normalize_job({
            "price_desc": "3박_다낭",
            "airline_code": "진에어",
            "hotel_name": "멜리아 빈펄 다낭",
            "progress_text": "예약마감",
            "rows": [{"date": "2026-07-15", "date_end": "2026-07-15"}],
            "source": "jobs.xlsx",
        })

        self.assertEqual(job["airline_code"], "LJ")
        self.assertEqual(job["hotel_name"], "멜리아 빈펄 다낭")
        self.assertEqual(job["rows"][0]["airline_code"], "LJ")
        self.assertIn("요금구분 : 3박_다낭 / 항공사 : LJ / 호텔명 : 멜리아 빈펄 다낭", job["rows"][0]["_job_label"])
        self.assertNotIn("진행구분", job["rows"][0]["_job_label"])

    def test_hotel_db_selection_keeps_hotel_seq_in_job(self):
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.hotel_name_var = Var("두짓")
        app.current_hotel_seq = ""
        app.hotel_choices = []
        app.hotel_record_by_label = {}
        app.hotel_record_by_seq = {}

        app._set_hotel_choices([{
            "infoSeq": 7864,
            "infoCd": "H",
            "infoTitle": "아미아나 리조트 나트랑",
            "natNm": "베트남",
            "cityNm": "나트랑",
            "useYn": "Y",
        }])
        app._select_hotel_record(app.hotel_choices[0])

        hotel_name, hotel_seq = app._selected_hotel_filter()
        job = app._normalize_job({
            "price_desc": "3박_나트랑",
            "airline_code": "7C",
            "hotel_name": hotel_name,
            "hotel_seq": hotel_seq,
            "rows": [{"date": "2026-07-15", "date_end": "2026-07-15"}],
            "source": "입력표",
        })

        self.assertEqual(hotel_name, "아미아나 리조트 나트랑")
        self.assertEqual(hotel_seq, "7864")
        self.assertEqual(job["hotel_seq"], "7864")
        self.assertEqual(job["rows"][0]["hotel_seq"], "7864")
        self.assertIn("호텔명 : 아미아나 리조트 나트랑 (hotelSeq 7864)", job["rows"][0]["_job_label"])

    def test_hotel_no_match_result_message_is_shown(self):
        class ListBox:
            def __init__(self):
                self.items = []
                self.item_options = {}

            def delete(self, *_args):
                self.items = []

            def insert(self, _index, value):
                self.items.append(value)

            def itemconfig(self, index, **kwargs):
                self.item_options[index] = kwargs

        class Frame:
            def __init__(self):
                self.mapped = False

            def winfo_ismapped(self):
                return self.mapped

            def pack(self, **_kwargs):
                self.mapped = True

            def pack_forget(self):
                self.mapped = False

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.hotel_result_list = ListBox()
        app.hotel_result_frame = Frame()
        app.hotel_result_records = []
        app.fg_muted = "#94a3b8"

        app._show_hotel_result_records([])

        self.assertEqual(app.hotel_result_list.items, ["검색 결과 없음"])
        self.assertTrue(app.hotel_result_frame.mapped)
        self.assertEqual(app.hotel_result_records, [])

    def test_direct_run_can_import_blank_conditions_from_erp_screen(self):
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        class Driver:
            def __init__(self):
                self.quit_called = False

            def quit(self):
                self.quit_called = True

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.config = {"selectors": {"search_date_input": "#searchStDate"}}
        app.driver = None
        app.price_desc_var = Var("")
        app.airline_var = Var(AIRLINE_EMPTY_LABEL)
        app.hotel_name_var = Var("")
        app.current_hotel_seq = ""
        app.hotel_choices = []
        app.hotel_record_by_label = {}
        app.hotel_record_by_seq = {}
        app._set_airline_choices([("LJ", "진에어"), ("KE", "대한항공")])
        driver = Driver()
        app._connect_matching_debug_browser = lambda *_args, **_kwargs: (driver, {"address": "127.0.0.1:9223"})
        app.find_and_switch_frame = lambda *_args, **_kwargs: True
        app._read_erp_screen_conditions = lambda _selectors: {
            "price_desc": "2N3D_TYO_LCC_ICN",
            "airline_code": "LJ",
            "airline_text": "[LJ] 진에어",
            "hotel_name": "아미아나 나트랑",
            "hotel_seq": "7864",
        }

        with mock.patch("gui.messagebox.askyesno", return_value=True) as askyesno:
            conditions = app._try_import_erp_conditions_for_direct_run()

        self.assertTrue(driver.quit_called)
        self.assertEqual(conditions["price_desc"], "2N3D_TYO_LCC_ICN")
        self.assertEqual(conditions["airline_code"], "LJ")
        self.assertEqual(conditions["hotel_name"], "아미아나 나트랑")
        self.assertEqual(conditions["hotel_seq"], "7864")
        self.assertEqual(app.price_desc_var.get(), "2N3D_TYO_LCC_ICN")
        self.assertEqual(app._selected_airline_code(), "LJ")
        self.assertEqual(app._selected_hotel_filter(), ("아미아나 나트랑", "7864"))
        confirm_msg = askyesno.call_args.args[1]
        self.assertIn("요금구분: 2N3D_TYO_LCC_ICN", confirm_msg)
        self.assertIn("항공사: [LJ] 진에어", confirm_msg)
        self.assertIn("호텔명: 아미아나 나트랑 (hotelSeq 7864)", confirm_msg)

    def test_erp_condition_import_keeps_existing_program_conditions(self):
        current = {
            "price_desc": "프로그램_요금",
            "airline_code": "KE",
            "hotel_name": "",
            "hotel_seq": "",
        }
        erp = {
            "price_desc": "ERP_요금",
            "airline_code": "LJ",
            "airline_text": "[LJ] 진에어",
            "hotel_name": "아미아나 나트랑",
            "hotel_seq": "7864",
        }

        pending = RpaGuiApp._pending_erp_condition_imports(current, erp)

        self.assertEqual(pending, {"hotel_name": "아미아나 나트랑", "hotel_seq": "7864"})

    def test_job_progress_status_counts_results(self):
        class Root:
            def after(self, _delay, func):
                func()

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.root = Root()
        app.job_queue = [{
            "price_desc": "3박_다낭",
            "airline_code": "LJ",
            "hotel_name": "",
            "rows": [{"date": "2026-07-15"}, {"date": "2026-07-16"}],
            "source": "입력표",
        }]
        app._refresh_job_queue_view = lambda: None
        app.job_queue = [app._normalize_job(app.job_queue[0])]
        app._assign_job_queue_metadata()

        app._set_job_progress_ui(0, status="진행 중")
        app._set_job_progress_ui(0, result_status="SUCCESS")
        app._set_job_progress_ui(0, result_status="SKIP")

        self.assertIn("완료", app._job_status_text(app.job_queue[0]))
        self.assertIn("성공 1", app._job_status_text(app.job_queue[0]))
        self.assertIn("건너뜀 1", app._job_status_text(app.job_queue[0]))

    def test_add_current_sheet_to_job_queue_accepts_hotel_only_condition(self):
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.is_running = False
        app.job_queue = []
        app.editing_job_index = None
        app.price_desc_var = Var("")
        app.hotel_name_var = Var("두짓비치 괌")
        app.read_sheet_data = lambda: ([{"date": "2026-07-15", "date_end": "2026-07-15"}], [])
        app._apply_date_filter = lambda rows: rows
        app._selected_airline_code = lambda: ""
        app._set_job_queue = lambda jobs, select_index=None: setattr(app, "job_queue", [app._normalize_job(job) for job in jobs])
        app._clear_sheet_after_job_registration = lambda: setattr(app, "sheet_cleared", True)
        app._select_job_index = lambda index: setattr(app, "selected_job_index", index)
        app.set_status = lambda *_args, **_kwargs: None
        app.accent_green = "#22c55e"

        with mock.patch("gui.messagebox.askyesno", return_value=True) as askyesno:
            app.add_current_sheet_to_job_queue()

        self.assertEqual(len(app.job_queue), 1)
        self.assertEqual(app.job_queue[0]["price_desc"], "")
        self.assertEqual(app.job_queue[0]["airline_code"], "")
        self.assertEqual(app.job_queue[0]["hotel_name"], "두짓비치 괌")
        self.assertTrue(app.sheet_cleared)
        confirm_msg = askyesno.call_args.args[1]
        self.assertIn("요금구분: 전체 요금구분", confirm_msg)
        self.assertIn("항공사: 전체 항공사", confirm_msg)
        self.assertIn("호텔명: 두짓비치 괌", confirm_msg)

    def test_add_current_sheet_to_job_queue_cancel_keeps_queue_unchanged(self):
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.is_running = False
        app.job_queue = []
        app.editing_job_index = None
        app.price_desc_var = Var("괌_저녁_3박")
        app.hotel_name_var = Var("")
        app.read_sheet_data = lambda: ([{"date": "2026-07-15", "date_end": "2026-07-15"}], [])
        app._apply_date_filter = lambda rows: rows
        app._selected_airline_code = lambda: ""
        app._set_job_queue = lambda *_args, **_kwargs: self.fail("queue should not be updated")
        app._clear_sheet_after_job_registration = lambda: setattr(app, "sheet_cleared", True)
        app._select_job_index = lambda *_args, **_kwargs: None
        app.set_status = lambda *_args, **_kwargs: None
        app.accent_green = "#22c55e"

        with mock.patch("gui.messagebox.askyesno", return_value=False):
            app.add_current_sheet_to_job_queue()

        self.assertEqual(app.job_queue, [])
        self.assertFalse(hasattr(app, "sheet_cleared"))

    def test_calculation_debug_groups_round_trips_by_night(self):
        result = calculate_round_trips(
            DEP_TEXT,
            RET_TEXT,
            FARES,
            SEASONS,
            "ICN-GUM-LJ",
            nights=(3, 4),
            today=date(2026, 6, 11),
        )
        app = RpaGuiApp.__new__(RpaGuiApp)

        text = "\n".join(app._format_round_trip_debug_by_night(result.debug.combinations))

        self.assertIn("[3박]", text)
        self.assertIn("[4박]", text)
        self.assertNotIn("출발일\t귀국일\t박수\t", text)
        self.assertIn("출발일\t귀국일\t출발편도\t귀국편도\t왕복\t상태\t시즌", text)

    def test_closed_exclusion_copy_text_lists_dates(self):
        app = RpaGuiApp.__new__(RpaGuiApp)
        rows = [
            RoundTripResult(
                dep_date="2026-07-01",
                ret_date="2026-07-04",
                nights=3,
                total_fare=0,
                is_closed=True,
                dep_fare=0,
                ret_fare=0,
            ),
            RoundTripResult(
                dep_date="2026-07-03",
                ret_date="2026-07-06",
                nights=3,
                total_fare=0,
                is_closed=True,
                dep_fare=0,
                ret_fare=0,
            ),
        ]

        text = app._closed_rows_copy_text(rows, 3)

        self.assertEqual(
            text.splitlines(),
            [
                "3박 마감 제외 목록",
                "출발일\t귀국일\t박수\t상태",
                "2026-07-01\t2026-07-04\t3\t마감",
                "2026-07-03\t2026-07-06\t3\t마감",
            ],
        )

    def test_js_round(self):
        self.assertEqual(js_round(100.5), 101)

    def test_excel_export(self):
        result = calculate_round_trips(
            DEP_TEXT,
            RET_TEXT,
            FARES,
            SEASONS,
            "ICN-GUM-LJ",
            nights=(3,),
            today=date(2026, 6, 11),
        )
        with tempfile.TemporaryDirectory() as tmp:
            path = export_results_to_excel(Path(tmp) / "out.xlsx", result.result, {"노선": "ICN-GUM-LJ"})
            self.assertTrue(path.exists())


if __name__ == "__main__":
    unittest.main()
