from __future__ import annotations

import argparse
import csv
import json
import sys
from pathlib import Path

from topas.availability import (
    availability_rows,
    build_ac1_workflow_commands,
    build_availability_commands,
    parse_availability_text,
)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="TOPAS availability helper")
    subparsers = parser.add_subparsers(dest="command", required=True)

    commands_parser = subparsers.add_parser("commands", help="build AN commands")
    commands_parser.add_argument("start_date")
    commands_parser.add_argument("end_date")
    commands_parser.add_argument("origin")
    commands_parser.add_argument("destination")
    commands_parser.add_argument("airline")
    commands_parser.add_argument("flight")
    commands_parser.add_argument(
        "--mode",
        choices=("ac1", "direct"),
        default="ac1",
        help="ac1: first AN then AC1 per day, direct: AN command for every date",
    )

    parse_parser = subparsers.add_parser("parse", help="parse TOPAS response text")
    parse_parser.add_argument("text_file")
    parse_parser.add_argument("--year", type=int, default=None)
    parse_parser.add_argument("--csv", dest="csv_path", default=None)

    args = parser.parse_args(argv)

    if args.command == "commands":
        builder = build_ac1_workflow_commands if args.mode == "ac1" else build_availability_commands
        for command in builder(
            args.start_date, args.end_date, args.origin, args.destination, args.airline, args.flight
        ):
            print(command)
        return 0

    if args.command == "parse":
        text_path = Path(args.text_file)
        text = text_path.read_text(encoding="utf-8")
        rows = availability_rows(parse_availability_text(text, year_hint=args.year))

        if args.csv_path:
            csv_path = Path(args.csv_path)
            fieldnames = sorted({key for row in rows for key in row.keys()})
            preferred = [
                "date",
                "weekday",
                "origin",
                "destination",
                "airline",
                "flight",
                "depart_time",
                "arrive_time",
                "equipment",
                "duration",
            ]
            fieldnames = preferred + [name for name in fieldnames if name not in preferred]
            with csv_path.open("w", encoding="utf-8-sig", newline="") as handle:
                writer = csv.DictWriter(handle, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
            print(csv_path)
            return 0

        print(json.dumps(rows, ensure_ascii=False, indent=2))
        return 0

    return 2


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
