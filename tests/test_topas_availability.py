from datetime import date
import unittest

from topas.availability import (
    availability_rows,
    build_ac1_workflow_commands,
    build_availability_commands,
    parse_availability_text,
)
from topas.pacing import TopasPacingPolicy


SAMPLE = """AN17AUGICNGUM/ALJ915 -AC-
** AMADEUS AVAILABILITY - AN ** GUM GUAM.GU        68 MO 17AUG 0000
 1   LJ 915   Y9 W9 D9 E9 H9 K9 L9 /ICN 2 GUM I 1815 2345 E0/738      4:30
       Q9 B9 N9 M9 X9 V9 Z9 A9 I9 T9
>"""

SELLCONNECT_SAMPLE = """>
AN01JULICNGUM/ALJ915
AN01JULICNGUM/ALJ915
** AMADEUS AVAILABILITY - AN ** GUM GUAM.GU                   21 WE 01JUL 0000
 1   LJ 915  Y9 W9 D9 E9 H9 K9 L9 /ICN 2 GUM I  1815    2345  E0/738       4:30
             Q9 B9 N9 M9 X9 P3 V9 Z9 A9 I9 O9 T9

>
AC1
AN02JULICNGUM/ALJ0915 -AC-
** AMADEUS AVAILABILITY - AN ** GUM GUAM.GU                   22 TH 02JUL 0000
 1   LJ 915  Y9 W9 D9 E9 H9 K9 L9 /ICN 2 GUM I  1815    2345  E0/738       4:30
             Q9 B9 N9 M9 X9 P9 S9 V9 Z9 A9 I9 O9 T9
>"""

NO_FLIGHT_SAMPLE = """>
AC1
AN03JULICNGUM/ALJ0915 -AC-
NO FLIGHT
>"""

DAD_SAMPLE = """AN1AUGICNDAD/AZE593
** AMADEUS AVAILABILITY - AN ** DAD DA NANG.VN              52 SA 01AUG 0000
 1   ZE 593  Y9 B9 H9 J9 K9 L9 M9 /ICN 1 DAD 2  2155    0040+1E0/738       4:45
             N9 P9 Q9 GR FR
>"""

PROMPT_ONLY_SAMPLE = """>
AC1
>
AC1
>
AC1
>"""

MARCH_WITH_PROMPT_NOISE = """>
AC1
>
AC1
>
AN21MARICNDAD/AZE0593 -AC-
** AMADEUS AVAILABILITY - AN ** DAD DA NANG.VN               284 SU 21MAR 0000
 1   ZE 593  Y9 B9 H9 J9 K9 L9 M9 /ICN 1 DAD 2  2055    0020+1E0/738       5:25
             N9 O9 P9 Q9 R9 S9 V9 W9 GR FR
>
AC1
AN22MARICNDAD/AZE0593 -AC-
** AMADEUS AVAILABILITY - AN ** DAD DA NANG.VN               285 MO 22MAR 0000
 1   ZE 593  Y9 B9 H9 J9 K9 L9 M9 /ICN 1 DAD 2  2055    0020+1E0/738       5:25
             N9 O9 P9 Q9 R9 S9 V9 W9 GR FR
>
AC1
AN28MARICNDAD/AZE0593 -AC-
** AMADEUS AVAILABILITY - AN ** DAD DA NANG.VN               291 SU 28MAR 0000
NO FLIGHT FOR THIS CITY PAIR - ENTER A CONNECT POINT /X...
CK ALT*ORIG GMP SSN XSM
>"""

AIRLINE_ONLY_ZE_MIXED_SAMPLE = """>
AN17SEPICNCNX/AZE -AC-
** AMADEUS AVAILABILITY - AN ** CNX CHIANG MAI.TH             98 TH 17SEP 0000
 1   ZE 517  Y9 B9 H9 K9 M9 /ICN 1 CNX I  1810    2205  E0/738       5:55
             N9 Q9 S9 V9 W9
>
AC1
AN18SEPICNCNX/AZE -AC-
NO FLIGHT FOR THIS CITY PAIR - ENTER A CONNECT POINT /X...
CK ALT*ORIG GMP SSN XSM
>
AC1
AN19SEPICNCNX/AZE -AC-
** AMADEUS AVAILABILITY - AN ** CNX CHIANG MAI.TH            100 SA 19SEP 0000
 1   ZE 517  Y9 B9 H9 K9 M9 /ICN 1 CNX I  1810    2205  E0/738       5:55
             N9 Q9 S9 V9 W9
>"""


class TopasAvailabilityTest(unittest.TestCase):
    def test_builds_independent_daily_commands(self):
        self.assertEqual(
            build_availability_commands("2026-07-01", "2026-07-03", "ICN", "GUM", "LJ", "0915"),
            [
                "AN01JULICNGUM/ALJ915",
                "AN02JULICNGUM/ALJ915",
                "AN03JULICNGUM/ALJ915",
            ],
        )

    def test_builds_ac1_workflow_commands(self):
        self.assertEqual(
            build_ac1_workflow_commands("2026-07-01", "2026-07-03", "ICN", "GUM", "LJ", "0915"),
            [
                "AN01JULICNGUM/ALJ915",
                "AC1",
                "AC1",
            ],
        )

    def test_parses_availability_response(self):
        blocks = parse_availability_text(SAMPLE, year_hint=2026)

        self.assertEqual(len(blocks), 1)
        block = blocks[0]
        self.assertEqual(block.travel_date, date(2026, 8, 17))
        self.assertEqual(block.offset_days, 68)
        self.assertEqual(block.weekday, "MO")
        self.assertEqual(block.origin, "ICN")
        self.assertEqual(block.destination, "GUM")

        flight = block.flights[0]
        self.assertEqual(flight.airline, "LJ")
        self.assertEqual(flight.flight, "915")
        self.assertEqual(flight.depart_time, "1815")
        self.assertEqual(flight.arrive_time, "2345")
        self.assertEqual(flight.equipment, "738")
        self.assertEqual(flight.classes["Y"], "9")
        self.assertEqual(flight.classes["T"], "9")

    def test_flattens_to_rows(self):
        rows = availability_rows(parse_availability_text(SAMPLE, year_hint=2026))
        self.assertEqual(rows[0]["date"], "2026-08-17")
        self.assertEqual(rows[0]["class_Y"], "9")
        self.assertEqual(rows[0]["class_T"], "9")

    def test_ignores_typed_command_line_in_sellconnect_output(self):
        blocks = parse_availability_text(SELLCONNECT_SAMPLE, year_hint=2026)
        rows = availability_rows(blocks)

        self.assertEqual([block.travel_date for block in blocks], [date(2026, 7, 1), date(2026, 7, 2)])
        self.assertEqual([row["date"] for row in rows], ["2026-07-01", "2026-07-02"])
        self.assertEqual(rows[1]["class_S"], "9")

    def test_prompt_ready_requires_prompt_without_loading_marker(self):
        policy = TopasPacingPolicy()
        self.assertTrue(policy.is_ready_for_next_command(SAMPLE))
        self.assertFalse(policy.is_ready_for_next_command("AC1\n||\n>"))

    def test_keeps_no_flight_response_as_raw_block(self):
        blocks = parse_availability_text(NO_FLIGHT_SAMPLE, year_hint=2026)

        self.assertEqual(len(blocks), 1)
        self.assertEqual(blocks[0].travel_date, date(2026, 7, 3))
        self.assertEqual(blocks[0].flights, ())
        self.assertIn("NO FLIGHT", blocks[0].raw_text)

    def test_parses_dad_availability_with_next_day_arrival(self):
        blocks = parse_availability_text(DAD_SAMPLE, year_hint=2026)

        self.assertEqual(len(blocks), 1)
        block = blocks[0]
        self.assertEqual(block.travel_date, date(2026, 8, 1))
        self.assertEqual(block.origin, "ICN")
        self.assertEqual(block.destination, "DAD")
        self.assertEqual(len(block.flights), 1)
        flight = block.flights[0]
        self.assertEqual(flight.airline, "ZE")
        self.assertEqual(flight.flight, "593")
        self.assertEqual(flight.arrive_time, "0040")
        self.assertEqual(flight.classes["Y"], "9")
        self.assertEqual(flight.classes["F"], "R")

    def test_ignores_prompt_only_ac1_noise(self):
        self.assertEqual(parse_availability_text(PROMPT_ONLY_SAMPLE, year_hint=2026), [])

    def test_keeps_only_real_responses_when_ac1_noise_is_interleaved(self):
        blocks = parse_availability_text(MARCH_WITH_PROMPT_NOISE, year_hint=2026)

        self.assertEqual([block.travel_date for block in blocks], [
            date(2026, 3, 21),
            date(2026, 3, 22),
            date(2026, 3, 28),
        ])
        self.assertEqual(len(blocks[0].flights), 1)
        self.assertEqual(len(blocks[1].flights), 1)
        self.assertEqual(blocks[2].flights, ())
        self.assertIn("NO FLIGHT", blocks[2].raw_text)

    def test_keeps_airline_only_request_no_flight_blocks(self):
        blocks = parse_availability_text(AIRLINE_ONLY_ZE_MIXED_SAMPLE, year_hint=2026)

        self.assertEqual([block.travel_date for block in blocks], [
            date(2026, 9, 17),
            date(2026, 9, 18),
            date(2026, 9, 19),
        ])
        self.assertEqual(blocks[0].request_command, "AN17SEPICNCNX/AZE -AC-")
        self.assertEqual(blocks[0].origin, "ICN")
        self.assertEqual(blocks[0].destination, "CNX")
        self.assertEqual(len(blocks[0].flights), 1)
        self.assertEqual(blocks[0].flights[0].airline, "ZE")
        self.assertEqual(blocks[1].request_command, "AN18SEPICNCNX/AZE -AC-")
        self.assertEqual(blocks[1].flights, ())
        self.assertIn("NO FLIGHT", blocks[1].raw_text)
        self.assertEqual(len(blocks[2].flights), 1)


if __name__ == "__main__":
    unittest.main()
