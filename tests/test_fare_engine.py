from datetime import date
from pathlib import Path
import tempfile
import unittest

from fare.calculator import RoundTripResult, calculate_round_trips, js_round
from fare.exporter import export_results_to_excel, results_to_tsv, to_erp_rows
from fare.parser import parse_topas_text
from gui import RpaGuiApp


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
