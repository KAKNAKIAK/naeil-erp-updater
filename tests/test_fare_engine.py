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
    def _make_fill_app(self, data, formulas=None, results=None):
        class Var:
            def __init__(self, value):
                self.value = value

            def get(self):
                return self.value

        class Sheet:
            def __init__(self, rows):
                self.data = [list(row) for row in rows]
                self._headers = ["날짜", "항공비", "호텔비", "지상비", "여행경비", "알선수익", "소아요금", "유아요금"]
                self.redrawn = False

            def get_sheet_data(self):
                return self.data

            def get_cell_data(self, r, c):
                try:
                    return self.data[r][c]
                except IndexError:
                    return ""

            def set_cell_data(self, r, c, value, **_kwargs):
                while len(self.data[r]) <= c:
                    self.data[r].append("")
                self.data[r][c] = value

            def headers(self):
                return self._headers

            def redraw(self):
                self.redrawn = True

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.period_mode_var = Var(False)
        app.sheet = Sheet(data)
        app.formulas = dict(formulas or {})
        app._results = dict(results or {})
        app._recalc_busy = False
        app._record_sheet_undo_state = lambda *_args, **_kwargs: None
        app.refresh_count = lambda: None
        app._load_active_into_fb = lambda: None
        app._sync_sheet_undo_baseline = lambda: None
        app.accent_green = "green"
        app.status = None
        app.set_status = lambda text, color: setattr(app, "status", (text, color))
        return app

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

    def test_progress_status_text_maps_all_supported_erp_codes(self):
        self.assertEqual(RpaGuiApp._progress_status_from_text("대기예약"), ("06", "대기예약"))
        self.assertEqual(RpaGuiApp._progress_status_from_text("예약 신청"), ("04", "예약신청"))
        self.assertEqual(RpaGuiApp._progress_status_from_text("예약 마감"), ("05", "예약마감"))

    def test_hotel_update_promotes_pending_reservation_before_price_update(self):
        self.assertTrue(
            RpaGuiApp._should_promote_pending_for_hotel_update(
                "06",
                {"adult_hotel_input": "434415"},
            )
        )
        self.assertTrue(
            RpaGuiApp._should_promote_pending_for_hotel_update(
                "06",
                {"adult_hotel_input": "0"},
            )
        )
        self.assertFalse(
            RpaGuiApp._should_promote_pending_for_hotel_update(
                "06",
                {"adult_air_input": "390000"},
            )
        )
        self.assertFalse(
            RpaGuiApp._should_promote_pending_for_hotel_update(
                "04",
                {"adult_hotel_input": "434415"},
            )
        )

    def test_sheet_pending_reservation_text_becomes_progress_status(self):
        class Var:
            def __init__(self, value):
                self.value = value

            def get(self):
                return self.value

        class Sheet:
            def get_sheet_data(self):
                return [["2026-07-15", "대기예약", "", "", "", "", "", ""]]

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.period_mode_var = Var(False)
        app.sheet = Sheet()

        rows, errors = app.read_sheet_data()

        self.assertEqual(errors, [])
        self.assertEqual(rows[0]["progress_status"], "대기예약")
        self.assertEqual(rows[0]["progress_status_field"], "adult_air")
        self.assertEqual(rows[0]["adult_air"], "")
        self.assertEqual(app._progress_status_from_text(rows[0]["progress_status"]), ("06", "대기예약"))

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

    def test_double_click_fill_applies_plain_value_to_last_data_row(self):
        app = self._make_fill_app(
            [
                ["2026-07-15", "예약마감", "", "", "", "", "", ""],
                ["2026-07-16", "100000", "", "", "", "", "", ""],
                ["2026-07-17", "", "", "", "", "", "", ""],
            ],
            formulas={(1, 1): "=호텔비"},
            results={(1, 1): "0"},
        )

        with mock.patch("gui.messagebox.askyesno", return_value=True) as askyesno:
            app._prompt_apply_cell_to_column(0, 1)

        askyesno.assert_called_once()
        self.assertEqual([row[1] for row in app.sheet.data], ["예약마감", "예약마감", "예약마감"])
        self.assertEqual(app.formulas, {})
        self.assertEqual(app._results, {})
        self.assertEqual(app.status, ("'항공비' 열 3개 행에 값을 채웠습니다.", "green"))

    def test_double_click_fill_keeps_formula_as_formula(self):
        app = self._make_fill_app(
            [
                ["2026-07-15", "100000", "20000", "", "", "", "", ""],
                ["2026-07-16", "300000", "40000", "", "", "", "", ""],
            ],
            formulas={(0, 5): "=항공비+호텔비"},
        )

        with mock.patch("gui.messagebox.askyesno", return_value=True):
            app._prompt_apply_cell_to_column(0, 5)

        self.assertEqual(app.formulas, {(0, 5): "=항공비+호텔비", (1, 5): "=항공비+호텔비"})
        self.assertEqual(app.sheet.data[0][5], "120000")
        self.assertEqual(app.sheet.data[1][5], "340000")
        self.assertEqual(app.status, ("'알선수익' 열 2개 행에 수식을 적용했습니다.", "green"))

    def test_current_page_progress_counts_detects_reservation_closed(self):
        class Driver:
            def execute_script(self, *_args):
                return [
                    {"procCd": "05", "procNm": "예약마감"},
                    {"procCd": "04", "procNm": "예약신청"},
                    {"procCd": "06", "procNm": "대기예약"},
                    {"procCd": "", "procNm": "예약 마감"},
                ]

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.config = {"grid_id": "#gridMain"}
        app.driver = Driver()

        summary = app._current_page_progress_status_counts()

        self.assertEqual(summary["total"], 4)
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
            "departure_flight": "LJ091",
            "hotel_name": "멜리아 빈펄 다낭",
            "progress_text": "예약마감",
            "rows": [{"date": "2026-07-15", "date_end": "2026-07-15"}],
            "source": "jobs.xlsx",
        })

        self.assertEqual(job["airline_code"], "LJ")
        self.assertEqual(job["departure_flight"], "LJ091")
        self.assertEqual(job["hotel_name"], "멜리아 빈펄 다낭")
        self.assertEqual(job["rows"][0]["airline_code"], "LJ")
        self.assertIn("요금구분 : 3박_다낭 / 항공사 : LJ / 출발편 : LJ091 / 호텔명 : 멜리아 빈펄 다낭", job["rows"][0]["_job_label"])
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
        _name, _seq, hotel_seq_strict = app._selected_hotel_filter_with_strict()
        job = app._normalize_job({
            "price_desc": "3박_나트랑",
            "airline_code": "7C",
            "hotel_name": hotel_name,
            "hotel_seq": hotel_seq,
            "hotel_seq_strict": hotel_seq_strict,
            "rows": [{"date": "2026-07-15", "date_end": "2026-07-15"}],
            "source": "입력표",
        })

        self.assertEqual(hotel_name, "아미아나 리조트 나트랑")
        self.assertEqual(hotel_seq, "7864")
        self.assertTrue(hotel_seq_strict)
        self.assertEqual(job["hotel_seq"], "7864")
        self.assertTrue(job["hotel_seq_strict"])
        self.assertEqual(job["rows"][0]["hotel_seq"], "7864")
        self.assertTrue(job["rows"][0]["hotel_seq_strict"])
        self.assertIn("호텔명 : 아미아나 리조트 나트랑 (hotelSeq 7864)", job["rows"][0]["_job_label"])

    def test_typed_hotel_name_uses_seq_as_reference_not_strict_validation(self):
        class Var:
            def __init__(self, value=""):
                self.value = value

            def get(self):
                return self.value

            def set(self, value):
                self.value = value

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.hotel_name_var = Var("괌 PIC 리조트")
        app.current_hotel_seq = "1796"
        app.current_hotel_seq_strict = False
        app.hotel_choices = []
        app.hotel_record_by_label = {}
        app.hotel_record_by_seq = {}
        app._set_hotel_choices([{
            "infoSeq": 1796,
            "infoCd": "H",
            "infoTitle": "괌 PIC 리조트",
            "natNm": "괌",
            "cityNm": "괌",
            "useYn": "Y",
        }])

        hotel_name, hotel_seq, hotel_seq_strict = app._selected_hotel_filter_with_strict()
        job = app._normalize_job({
            "price_desc": "김괌",
            "airline_code": "",
            "hotel_name": hotel_name,
            "hotel_seq": hotel_seq,
            "hotel_seq_strict": hotel_seq_strict,
            "rows": [{"date": "2026-07-01", "date_end": "2026-07-23"}],
            "source": "입력표",
        })

        self.assertEqual((hotel_name, hotel_seq, hotel_seq_strict), ("괌 PIC 리조트", "1796", False))
        self.assertEqual(job["hotel_seq"], "1796")
        self.assertFalse(job["hotel_seq_strict"])
        self.assertFalse(job["rows"][0]["hotel_seq_strict"])
        self.assertIn("호텔명 : 괌 PIC 리조트", job["rows"][0]["_job_label"])
        self.assertNotIn("hotelSeq 1796", job["rows"][0]["_job_label"])

    def test_selected_hotel_allows_different_erp_seq_when_name_matches(self):
        class Driver:
            def execute_script(self, script, *args):
                if "const nameSelector = arguments[0]" in script:
                    return {"ok": True, "value": "아미아나 리조트 나트랑", "seq": "9001", "cleared": False}
                raise AssertionError(f"Unexpected script: {script[:80]}")

        app = RpaGuiApp.__new__(RpaGuiApp)
        app.driver = Driver()

        result = app._set_erp_hotel_filter(
            {"hotel_name_input": "#hotelKorNm", "hotel_seq_input": "#hotelSeq"},
            "아미아나 리조트 나트랑",
            expected_seq="7864",
        )

        self.assertEqual(result["value"], "아미아나 리조트 나트랑")
        self.assertEqual(result["seq"], "9001")
        self.assertEqual(result["selected_seq"], "7864")

    def test_hotel_filter_clears_stale_seq_before_waiting_for_new_seq(self):
        class Driver:
            def __init__(self):
                self.setup_script = ""

            def execute_script(self, script, *args):
                if "const nameSelector = arguments[0]" in script:
                    self.setup_script = script
                    return {"ok": True, "value": "괌 니코 호텔", "seq": "", "previousSeq": "2264", "cleared": False}
                if "const seq = document.querySelector" in script:
                    return "4610"
                if "const el = document.querySelector" in script:
                    return "괌 니코 호텔"
                raise AssertionError(f"Unexpected script: {script[:80]}")

        app = RpaGuiApp.__new__(RpaGuiApp)
        driver = Driver()
        app.driver = driver

        with mock.patch("gui.time.sleep", return_value=None):
            result = app._set_erp_hotel_filter(
                {"hotel_name_input": "#hotelKorNm", "hotel_seq_input": "#hotelSeq"},
                "괌 니코 호텔",
                expected_seq="4610",
                timeout=0.2,
                poll=0.01,
            )

        self.assertIn("seq.value = ''", driver.setup_script)
        self.assertEqual(result["value"], "괌 니코 호텔")
        self.assertEqual(result["seq"], "4610")

    def test_hotel_name_match_ignores_spacing(self):
        self.assertTrue(RpaGuiApp._hotel_name_matches("괌PIC리조트", "괌 PIC 리조트"))
        self.assertFalse(RpaGuiApp._hotel_name_matches("괌 니코 호텔", "괌 PIC 리조트"))

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
        app.departure_flight_var = Var("")
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
            "departure_flight": "LJ091",
            "airline_code": "LJ",
            "airline_text": "[LJ] 진에어",
            "hotel_name": "아미아나 나트랑",
            "hotel_seq": "7864",
        }

        with mock.patch("gui.messagebox.askyesnocancel", return_value=True) as askyesnocancel:
            conditions = app._try_import_erp_conditions_for_direct_run()

        self.assertTrue(driver.quit_called)
        self.assertEqual(conditions["price_desc"], "2N3D_TYO_LCC_ICN")
        self.assertEqual(conditions["departure_flight"], "LJ091")
        self.assertEqual(conditions["airline_code"], "LJ")
        self.assertEqual(conditions["hotel_name"], "아미아나 나트랑")
        self.assertEqual(conditions["hotel_seq"], "7864")
        self.assertFalse(conditions["hotel_seq_strict"])
        self.assertEqual(app.price_desc_var.get(), "2N3D_TYO_LCC_ICN")
        self.assertEqual(app.departure_flight_var.get(), "LJ091")
        self.assertEqual(app._selected_airline_code(), "LJ")
        self.assertEqual(app._selected_hotel_filter(), ("아미아나 나트랑", "7864"))
        confirm_msg = askyesnocancel.call_args.args[1]
        self.assertIn("요금구분: 2N3D_TYO_LCC_ICN", confirm_msg)
        self.assertIn("출발편: LJ091", confirm_msg)
        self.assertIn("항공사: [LJ] 진에어", confirm_msg)
        self.assertIn("호텔명: 아미아나 나트랑 (hotelSeq 7864)", confirm_msg)
        self.assertIn("취소: 실행하지 않고 입력표로 돌아가기", confirm_msg)

    def test_direct_run_condition_import_cancel_aborts_run_choice(self):
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
        app.departure_flight_var = Var("")
        app.airline_var = Var(AIRLINE_EMPTY_LABEL)
        app.hotel_name_var = Var("")
        app.current_hotel_seq = ""
        app.hotel_choices = []
        app.hotel_record_by_label = {}
        app.hotel_record_by_seq = {}
        app._set_airline_choices([("LJ", "진에어")])
        driver = Driver()
        app._connect_matching_debug_browser = lambda *_args, **_kwargs: (driver, {"address": "127.0.0.1:9223"})
        app.find_and_switch_frame = lambda *_args, **_kwargs: True
        app._read_erp_screen_conditions = lambda _selectors: {
            "price_desc": "ERP_요금",
            "departure_flight": "ZE581",
            "airline_code": "LJ",
            "airline_text": "[LJ] 진에어",
            "hotel_name": "",
            "hotel_seq": "",
        }

        with mock.patch("gui.messagebox.askyesnocancel", return_value=None):
            conditions = app._try_import_erp_conditions_for_direct_run()

        self.assertTrue(driver.quit_called)
        self.assertIsNone(conditions)
        self.assertEqual(app.price_desc_var.get(), "")
        self.assertEqual(app.departure_flight_var.get(), "")
        self.assertEqual(app._selected_airline_code(), "")

    def test_erp_condition_import_keeps_existing_program_conditions(self):
        current = {
            "price_desc": "프로그램_요금",
            "departure_flight": "",
            "airline_code": "KE",
            "hotel_name": "",
            "hotel_seq": "",
        }
        erp = {
            "price_desc": "ERP_요금",
            "departure_flight": "ZE581",
            "airline_code": "LJ",
            "airline_text": "[LJ] 진에어",
            "hotel_name": "아미아나 나트랑",
            "hotel_seq": "7864",
        }

        pending = RpaGuiApp._pending_erp_condition_imports(current, erp)

        self.assertEqual(pending, {"departure_flight": "ZE581", "hotel_name": "아미아나 나트랑", "hotel_seq": "7864", "hotel_seq_strict": False})

    def test_job_queue_row_conditions_do_not_fallback_to_first_job_hotel_seq(self):
        app = RpaGuiApp.__new__(RpaGuiApp)
        app.selected_price_desc = "김괌"
        app.selected_airline_code = ""
        app.selected_departure_flight = "LJ091"
        app.selected_hotel_name = "괌 플라자 호텔"
        app.selected_hotel_seq = "6195"
        app.selected_hotel_seq_strict = True
        app.selected_progress_text = ""

        job_row = {
            "_job_index": 2,
            "price_desc": "김괌",
            "airline_code": "",
            "departure_flight": "LJ092",
            "hotel_name": "괌 니코 호텔",
            "hotel_seq": "",
            "hotel_seq_strict": False,
            "progress_status": "",
        }
        direct_row = {
            "price_desc": "",
            "airline_code": "",
            "departure_flight": "",
            "hotel_name": "",
            "hotel_seq": "",
            "hotel_seq_strict": "",
            "progress_status": "",
        }

        job_conditions = app._rpa_row_conditions(job_row)
        direct_conditions = app._rpa_row_conditions(direct_row)

        self.assertEqual(job_conditions["hotel_name"], "괌 니코 호텔")
        self.assertEqual(job_conditions["hotel_seq"], "")
        self.assertFalse(job_conditions["hotel_seq_strict"])
        self.assertEqual(job_conditions["price_desc"], "김괌")
        self.assertEqual(job_conditions["departure_flight"], "LJ092")
        self.assertEqual(direct_conditions["departure_flight"], "LJ091")
        self.assertEqual(direct_conditions["hotel_name"], "괌 플라자 호텔")
        self.assertEqual(direct_conditions["hotel_seq"], "6195")
        self.assertTrue(direct_conditions["hotel_seq_strict"])

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
