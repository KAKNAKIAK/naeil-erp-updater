import unittest

from fare.route_select import filter_routes, infer_route_from_topas_text


ROUTES = [
    "PUS-ICN-KE",
    "ICN-GUM-LJ",
    "ICN-CNX-ZE",
    "ICN-DAD-ZE",
    "ICN-BKK-TG",
    "GMP-CJU-7C",
]


DEP_CNX = """AN17SEPICNCNX/AZE -AC-
** AMADEUS AVAILABILITY - AN ** CNX CHIANG MAI.TH             98 TH 17SEP 0000
 1   ZE 517  Y9 B9 H9 K9 M9 /ICN 1 CNX I  1810    2205  E0/738       5:55
>"""


RET_CNX = """AN20SEPCNXICN/AZE -AC-
** AMADEUS AVAILABILITY - AN ** SEL SEOUL.KR                 101 SU 20SEP 0000
 1   ZE 518  Y9 B9 H9 K9 M9 /CNX I ICN 1  2305    0630+1E0/738       5:25
>"""


NO_FLIGHT_DEP = """AN18SEPICNCNX/AZE -AC-
NO FLIGHT FOR THIS CITY PAIR - ENTER A CONNECT POINT /X...
CK ALT*ORIG GMP SSN XSM
>"""


class RouteSelectTest(unittest.TestCase):
    def test_filters_routes_without_hyphen_or_case(self):
        self.assertEqual(filter_routes(ROUTES, "icncnx"), ["ICN-CNX-ZE"])
        self.assertEqual(filter_routes(ROUTES, "cnx ze"), ["ICN-CNX-ZE"])
        self.assertEqual(filter_routes(ROUTES, "gum"), ["ICN-GUM-LJ"])

    def test_prioritizes_airport_code_prefix_for_short_queries(self):
        matches = filter_routes(ROUTES, "I", limit=None)

        self.assertEqual(
            matches[:4],
            ["ICN-GUM-LJ", "ICN-CNX-ZE", "ICN-DAD-ZE", "ICN-BKK-TG"],
        )
        self.assertGreater(matches.index("PUS-ICN-KE"), matches.index("ICN-BKK-TG"))

    def test_searches_airport_and_airline_code_prefixes(self):
        self.assertEqual(filter_routes(ROUTES, "CN")[0], "ICN-CNX-ZE")
        self.assertEqual(filter_routes(ROUTES, "ZE"), ["ICN-CNX-ZE", "ICN-DAD-ZE"])

    def test_infers_route_from_departure_text(self):
        inference = infer_route_from_topas_text(DEP_CNX, "", ROUTES, year_hint=2026)

        self.assertEqual(inference.route, "ICN-CNX-ZE")
        self.assertIn("출발편", inference.reason)

    def test_infers_route_from_return_text_reversed(self):
        inference = infer_route_from_topas_text("", RET_CNX, ROUTES, year_hint=2026)

        self.assertEqual(inference.route, "ICN-CNX-ZE")
        self.assertIn("귀국편", inference.reason)

    def test_infers_route_from_airline_only_no_flight_text(self):
        inference = infer_route_from_topas_text(NO_FLIGHT_DEP, "", ROUTES, year_hint=2026)

        self.assertEqual(inference.route, "ICN-CNX-ZE")


if __name__ == "__main__":
    unittest.main()
