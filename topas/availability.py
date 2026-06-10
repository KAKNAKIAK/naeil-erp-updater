"""Pure TOPAS availability command generation and text parsing.

This module intentionally has no Selenium/tkinter dependency. The v4 TOPAS
feature can be tested here first, then wired into the GUI later.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
import re
from typing import Iterable


MONTHS = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}

MONTH_NAMES = {v: k for k, v in MONTHS.items()}

REQUEST_RE = re.compile(
    r"^AN(?P<day>\d{1,2})(?P<month>[A-Z]{3})"
    r"(?P<origin>[A-Z]{3})(?P<destination>[A-Z]{3})"
    r"/A(?P<airline>[A-Z0-9]{2})(?P<flight>\d{1,4})",
    re.IGNORECASE,
)

HEADER_RE = re.compile(
    r"\b(?P<offset>\d{1,3})\s+(?P<weekday>[A-Z]{2})\s+"
    r"(?P<day>\d{2})(?P<month>[A-Z]{3})\s+\d{4}\b",
    re.IGNORECASE,
)

FLIGHT_LINE_RE = re.compile(
    r"^\s*(?P<line_no>\d+)\s+"
    r"(?P<airline>[A-Z0-9]{2})\s+"
    r"(?P<flight>\d{1,4})\s+"
    r"(?P<class_text>.*?)\s+/"
    r"(?P<origin>[A-Z]{3})\s+"
    r"(?P<origin_terminal>\d+)?\s*"
    r"(?P<destination>[A-Z]{3})\s+"
    r"(?P<destination_terminal>[A-Z0-9]+)?\s+"
    r"(?P<depart_time>\d{4})\s+"
    r"(?P<arrive_time>\d{4})(?P<arrive_day_offset>[+-]\d+)?\s*"
    r"(?P<meal>[A-Z0-9]+)?/"
    r"(?P<equipment>[A-Z0-9]{3})\s+"
    r"(?P<duration>\d{1,2}:\d{2})",
    re.IGNORECASE,
)

CLASS_RE = re.compile(r"\b([A-Z])([0-9A-Z])\b", re.IGNORECASE)


@dataclass(frozen=True)
class FlightAvailability:
    line_no: int
    airline: str
    flight: str
    origin: str
    destination: str
    depart_time: str
    arrive_time: str
    equipment: str
    duration: str
    classes: dict[str, str] = field(default_factory=dict)
    raw_lines: tuple[str, ...] = field(default_factory=tuple)


@dataclass(frozen=True)
class AvailabilityBlock:
    request_command: str | None
    travel_date: date | None
    offset_days: int | None
    weekday: str | None
    origin: str | None
    destination: str | None
    flights: tuple[FlightAvailability, ...]
    raw_text: str


def parse_yyyymmdd(value: str | date) -> date:
    if isinstance(value, date):
        return value
    text = str(value).strip()
    for fmt in ("%Y-%m-%d", "%Y%m%d", "%Y/%m/%d"):
        try:
            return datetime.strptime(text, fmt).date()
        except ValueError:
            pass
    raise ValueError(f"Unsupported date format: {value!r}")


def format_topas_day_month(value: date) -> str:
    return f"{value.day:02d}{MONTH_NAMES[value.month]}"


def build_availability_commands(
    start_date: str | date,
    end_date: str | date,
    origin: str,
    destination: str,
    airline: str,
    flight: str | int,
) -> list[str]:
    """Build independent AN commands for every date in the inclusive range."""

    start = parse_yyyymmdd(start_date)
    end = parse_yyyymmdd(end_date)
    if end < start:
        raise ValueError("end_date must be on or after start_date")

    origin = origin.strip().upper()
    destination = destination.strip().upper()
    airline = airline.strip().upper()
    flight_text = str(flight).strip().upper().lstrip("0") or "0"

    commands: list[str] = []
    current = start
    while current <= end:
        commands.append(
            f"AN{format_topas_day_month(current)}{origin}{destination}/A{airline}{flight_text}"
        )
        current += timedelta(days=1)
    return commands


def build_initial_availability_command(
    travel_date: str | date,
    origin: str,
    destination: str,
    airline: str,
    flight: str | int,
) -> str:
    """Build the first absolute AN command for a SellConnect availability run."""

    return build_availability_commands(
        travel_date,
        travel_date,
        origin,
        destination,
        airline,
        flight,
    )[0]


def build_ac1_workflow_commands(
    start_date: str | date,
    end_date: str | date,
    origin: str,
    destination: str,
    airline: str,
    flight: str | int,
) -> list[str]:
    """Build the preferred sequential workflow: first AN, then AC1 per next day."""

    start = parse_yyyymmdd(start_date)
    end = parse_yyyymmdd(end_date)
    if end < start:
        raise ValueError("end_date must be on or after start_date")

    days = (end - start).days + 1
    first = build_initial_availability_command(start, origin, destination, airline, flight)
    return [first] + ["AC1"] * (days - 1)


def parse_availability_text(text: str, year_hint: int | None = None) -> list[AvailabilityBlock]:
    """Parse one or more TOPAS availability responses from terminal text."""

    blocks = _split_blocks(text)
    return [_parse_block(block, year_hint=year_hint) for block in blocks]


def _split_blocks(text: str) -> list[str]:
    lines = [line.rstrip() for line in text.replace("\r\n", "\n").split("\n")]
    request_lines = [idx for idx, line in enumerate(lines) if REQUEST_RE.search(line.strip())]
    if not request_lines:
        return []

    starts = request_lines + [len(lines)]
    blocks: list[str] = []
    for idx in range(len(starts) - 1):
        block_lines = lines[starts[idx] : starts[idx + 1]]
        block = "\n".join(block_lines).strip()
        block_upper = block.upper()
        useful_lines = [line.strip() for line in block_lines if line.strip() and line.strip() != ">"]
        # SellConnect keeps the typed command line above the echoed TOPAS
        # response. A single request line is only the typed command, while a
        # multi-line block is a TOPAS response even when it says NO FLIGHT.
        if "AMADEUS AVAILABILITY" not in block_upper and len(useful_lines) < 2:
            continue
        if block:
            blocks.append(block)
    return blocks


def _parse_block(block: str, year_hint: int | None) -> AvailabilityBlock:
    lines = [line.rstrip() for line in block.splitlines()]
    request_line = next((line.strip() for line in lines if REQUEST_RE.search(line.strip())), None)
    request = REQUEST_RE.search(request_line or "")
    header = next((HEADER_RE.search(line) for line in lines if HEADER_RE.search(line)), None)

    year = year_hint or date.today().year
    travel_date = None
    if header:
        travel_date = _resolve_day_month(header.group("day"), header.group("month"), year)
    elif request:
        travel_date = _resolve_day_month(request.group("day"), request.group("month"), year)

    flights: list[FlightAvailability] = []
    current_idx: int | None = None
    current_lines: list[str] = []

    def flush_current() -> None:
        nonlocal current_idx, current_lines
        if current_idx is None:
            return
        existing = flights[current_idx]
        class_text = " ".join(current_lines)
        merged_classes = dict(existing.classes)
        merged_classes.update(_parse_classes(class_text))
        flights[current_idx] = FlightAvailability(
            line_no=existing.line_no,
            airline=existing.airline,
            flight=existing.flight,
            origin=existing.origin,
            destination=existing.destination,
            depart_time=existing.depart_time,
            arrive_time=existing.arrive_time,
            equipment=existing.equipment,
            duration=existing.duration,
            classes=merged_classes,
            raw_lines=tuple(current_lines),
        )
        current_idx = None
        current_lines = []

    for line in lines:
        flight_match = FLIGHT_LINE_RE.match(line)
        if flight_match:
            flush_current()
            class_text = flight_match.group("class_text")
            flight = FlightAvailability(
                line_no=int(flight_match.group("line_no")),
                airline=flight_match.group("airline").upper(),
                flight=flight_match.group("flight"),
                origin=flight_match.group("origin").upper(),
                destination=flight_match.group("destination").upper(),
                depart_time=flight_match.group("depart_time"),
                arrive_time=flight_match.group("arrive_time"),
                equipment=flight_match.group("equipment").upper(),
                duration=flight_match.group("duration"),
                classes=_parse_classes(class_text),
                raw_lines=(line,),
            )
            flights.append(flight)
            current_idx = len(flights) - 1
            current_lines = [line]
            continue

        if current_idx is not None and _looks_like_class_continuation(line):
            current_lines.append(line)

    flush_current()

    origin = request.group("origin").upper() if request else (flights[0].origin if flights else None)
    destination = request.group("destination").upper() if request else (
        flights[0].destination if flights else None
    )

    return AvailabilityBlock(
        request_command=request_line,
        travel_date=travel_date,
        offset_days=int(header.group("offset")) if header else None,
        weekday=header.group("weekday").upper() if header else None,
        origin=origin,
        destination=destination,
        flights=tuple(flights),
        raw_text=block,
    )


def _resolve_day_month(day: str, month: str, year: int) -> date:
    return date(year, MONTHS[month.upper()], int(day))


def _parse_classes(text: str) -> dict[str, str]:
    return {code.upper(): value.upper() for code, value in CLASS_RE.findall(text)}


def _looks_like_class_continuation(line: str) -> bool:
    stripped = line.strip()
    return bool(stripped) and bool(CLASS_RE.search(stripped)) and "/" not in stripped


def availability_rows(blocks: Iterable[AvailabilityBlock]) -> list[dict[str, object]]:
    """Flatten parsed blocks into spreadsheet-friendly dictionaries."""

    rows: list[dict[str, object]] = []
    for block in blocks:
        for flight in block.flights:
            row: dict[str, object] = {
                "date": block.travel_date.isoformat() if block.travel_date else "",
                "weekday": block.weekday or "",
                "origin": flight.origin,
                "destination": flight.destination,
                "airline": flight.airline,
                "flight": flight.flight,
                "depart_time": flight.depart_time,
                "arrive_time": flight.arrive_time,
                "equipment": flight.equipment,
                "duration": flight.duration,
            }
            row.update({f"class_{code}": value for code, value in sorted(flight.classes.items())})
            rows.append(row)
    return rows
