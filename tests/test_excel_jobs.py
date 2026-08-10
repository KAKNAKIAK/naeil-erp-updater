import tempfile
from pathlib import Path
import unittest

import pandas as pd

import excel_loader


class ExcelJobImportTest(unittest.TestCase):
    def _write_excel(self, rows):
        tmp = tempfile.TemporaryDirectory()
        path = Path(tmp.name) / "jobs.xlsx"
        pd.DataFrame(rows).to_excel(path, index=False)
        self.addCleanup(tmp.cleanup)
        return path

    def test_condition_columns_build_jobs_by_price_and_airline(self):
        path = self._write_excel([
            {
                "요금구분": "3박_다낭",
                "항공사코드": "LJ",
                "출발편": "LJ091",
                "호텔명": "멜리아 빈펄 다낭",
                "진행구분": "예약마감",
                "시작일": "2026-07-15",
                "종료일": "2026-07-15",
                "항공비": 350000,
            },
            {
                "요금구분": "3박_다낭",
                "항공사코드": "LJ",
                "출발편": "LJ091",
                "호텔명": "멜리아 빈펄 다낭",
                "진행구분": "예약마감",
                "시작일": "2026-07-16",
                "종료일": "2026-07-16",
                "항공비": 360000,
            },
            {
                "요금구분": "3박_나트랑",
                "항공사코드": "7C",
                "출발편": "7C2901",
                "호텔명": "아미아나 리조트 나트랑",
                "진행구분": "예약마감",
                "시작일": "2026-07-17",
                "종료일": "2026-07-17",
                "항공비": 410000,
            },
        ])

        result = excel_loader.load_fare_jobs_from_excel(path)

        self.assertTrue(result["detected"])
        self.assertTrue(result["is_period"])
        self.assertEqual(len(result["jobs"]), 2)
        self.assertEqual(result["jobs"][0]["price_desc"], "3박_다낭")
        self.assertEqual(result["jobs"][0]["airline_code"], "LJ")
        self.assertEqual(result["jobs"][0]["departure_flight"], "LJ091")
        self.assertEqual(result["jobs"][0]["hotel_name"], "멜리아 빈펄 다낭")
        self.assertEqual(len(result["jobs"][0]["rows"]), 2)
        self.assertEqual(result["jobs"][1]["price_desc"], "3박_나트랑")
        self.assertEqual(result["jobs"][1]["departure_flight"], "7C2901")
        self.assertEqual(result["jobs"][1]["hotel_name"], "아미아나 리조트 나트랑")
        self.assertEqual(result["jobs"][1]["rows"][0]["adult_air"], 410000)
        self.assertEqual(result["jobs"][1]["rows"][0]["departure_flight"], "7C2901")
        self.assertEqual(result["jobs"][1]["rows"][0]["hotel_name"], "아미아나 리조트 나트랑")

    def test_reservation_closed_fare_cell_keeps_source_field(self):
        path = self._write_excel([
            {
                "요금구분": "괌_저녁_3박",
                "항공사코드": "LJ",
                "시작일": "2026-07-15",
                "항공비": "예약마감",
            },
        ])

        result = excel_loader.load_fare_jobs_from_excel(path)

        row = result["jobs"][0]["rows"][0]
        self.assertEqual(row["progress_status"], "예약마감")
        self.assertEqual(row["progress_status_field"], "adult_air")
        self.assertEqual(row["adult_air"], "")

    def test_pending_reservation_fare_cell_keeps_source_field(self):
        path = self._write_excel([
            {
                "요금구분": "괌_저녁_3박",
                "항공사코드": "LJ",
                "시작일": "2026-07-15",
                "항공비": "예약대기",
            },
        ])

        result = excel_loader.load_fare_jobs_from_excel(path)

        row = result["jobs"][0]["rows"][0]
        self.assertEqual(row["progress_status"], "대기예약")
        self.assertEqual(row["progress_status_field"], "adult_air")
        self.assertEqual(row["adult_air"], "")

    def test_legacy_excel_is_not_detected_as_job_excel(self):
        path = self._write_excel([
            {"날짜": "2026-07-15", "항공비": 350000, "알선수익": 0},
        ])

        result = excel_loader.load_fare_jobs_from_excel(path)

        self.assertFalse(result["detected"])
        self.assertEqual(result["jobs"], [])


if __name__ == "__main__":
    unittest.main()
