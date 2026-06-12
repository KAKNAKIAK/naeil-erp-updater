from datetime import date
import unittest

from gui import RpaGuiApp


def fare_record(start, end=None, fare="100000"):
    end = end or start
    fares = [fare, "", "", "", "", "", ""]
    return {
        "start": start,
        "end": end,
        "fares": fares,
        "key": tuple(fares),
    }


class SheetMergeTest(unittest.TestCase):
    def test_does_not_merge_across_missing_daily_date(self):
        records = [
            fare_record(date(2026, 7, 1)),
            fare_record(date(2026, 7, 2)),
            fare_record(date(2026, 7, 4)),
        ]

        merged, gaps = RpaGuiApp._merge_fare_records_preserving_gaps(records)

        self.assertEqual(
            [(row["start"], row["end"]) for row in merged],
            [
                (date(2026, 7, 1), date(2026, 7, 2)),
                (date(2026, 7, 4), date(2026, 7, 4)),
            ],
        )
        self.assertEqual(gaps, [(date(2026, 7, 3), date(2026, 7, 3))])

    def test_merges_adjacent_periods_only_when_there_is_no_gap(self):
        records = [
            fare_record(date(2026, 7, 1), date(2026, 7, 3)),
            fare_record(date(2026, 7, 4), date(2026, 7, 6)),
            fare_record(date(2026, 7, 8), date(2026, 7, 9)),
        ]

        merged, gaps = RpaGuiApp._merge_fare_records_preserving_gaps(records)

        self.assertEqual(
            [(row["start"], row["end"]) for row in merged],
            [
                (date(2026, 7, 1), date(2026, 7, 6)),
                (date(2026, 7, 8), date(2026, 7, 9)),
            ],
        )
        self.assertEqual(gaps, [(date(2026, 7, 7), date(2026, 7, 7))])


if __name__ == "__main__":
    unittest.main()
