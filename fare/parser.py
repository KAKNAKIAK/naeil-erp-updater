"""TOPAS raw text parser compatible with the deployed web fare calculator.

The web bundle parses AN blocks by command echo, resolves day/month against
today, and accepts only class tokens with 4-9 available seats. This module keeps
that behavior intentionally small and testable.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
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

AN_RE = re.compile(r"AN(\d{1,2}[A-Z]{3})", re.IGNORECASE)
REQUEST_ROUTE_RE = re.compile(
    r"AN\d{1,2}[A-Z]{3}(?P<origin>[A-Z]{3})(?P<destination>[A-Z]{3})/A",
    re.IGNORECASE,
)
CLASS_RE = re.compile(r"(?:^|\s+)([A-Z][4-9])(?!\S)", re.IGNORECASE)


@dataclass(frozen=True)
class ParsedAvailabilityRecord:
    date: str
    classes: tuple[str, ...]
    raw_text: str = ""
    origin: str | None = None
    destination: str | None = None


@dataclass(frozen=True)
class ParseResult:
    records: tuple[ParsedAvailabilityRecord, ...]
    warnings: tuple[str, ...] = field(default_factory=tuple)

    def by_date(self) -> dict[str, ParsedAvailabilityRecord]:
        return {record.date: record for record in self.records}


def parse_topas_text(text: str, today: date | None = None) -> ParseResult:
    return parse_topas_lines(str(text or "").replace("\r\n", "\n").splitlines(), today=today)


def parse_topas_lines(lines: Iterable[str], today: date | None = None) -> ParseResult:
    today = today or date.today()
    current_date: str | None = None
    current_lines: list[str] = []
    current_classes: list[str] = []
    current_origin: str | None = None
    current_destination: str | None = None
    ordered_dates: list[str] = []
    by_date: dict[str, ParsedAvailabilityRecord] = {}
    duplicate_dates: set[str] = set()
    warnings: list[str] = []

    def flush() -> None:
        nonlocal current_date, current_lines, current_classes, current_origin, current_destination
        if not current_date:
            return
        if current_date in by_date:
            duplicate_dates.add(current_date)
        else:
            ordered_dates.append(current_date)
        by_date[current_date] = ParsedAvailabilityRecord(
            date=current_date,
            classes=tuple(current_classes),
            raw_text="\n".join(current_lines).strip(),
            origin=current_origin,
            destination=current_destination,
        )
        current_date = None
        current_lines = []
        current_classes = []
        current_origin = None
        current_destination = None

    for raw_line in lines:
        line = str(raw_line).rstrip()
        command_match = AN_RE.search(line.strip())
        if command_match:
            flush()
            resolved = resolve_day_month(command_match.group(1).upper(), today=today)
            if resolved is None:
                continue
            current_date = resolved
            current_lines = [line]
            current_classes = []
            route_match = REQUEST_ROUTE_RE.search(line.strip())
            if route_match:
                current_origin = route_match.group("origin").upper()
                current_destination = route_match.group("destination").upper()
            continue

        if current_date:
            current_lines.append(line)
            tokens = [token.strip().upper() for token in CLASS_RE.findall(line)]
            current_classes.extend(token for token in tokens if token != "AC1")

    flush()

    if duplicate_dates:
        warnings.append(
            "중복 날짜 블록은 마지막 응답을 사용했습니다: "
            + ", ".join(sorted(duplicate_dates))
        )

    records = tuple(by_date[day] for day in ordered_dates if day in by_date)
    if not records:
        warnings.append("원문에서 조회 날짜를 찾지 못했습니다.")

    return ParseResult(records=records, warnings=tuple(warnings))


def resolve_day_month(day_month: str, today: date | None = None) -> str | None:
    today = today or date.today()
    text = str(day_month).strip().upper()
    day_text = text[:-3].zfill(2)
    month_text = text[-3:]
    month = MONTHS.get(month_text)
    if not month:
        return None
    try:
        candidate = date(today.year, month, int(day_text))
    except ValueError:
        return None
    if candidate.isoformat() < today.isoformat():
        candidate = date(today.year + 1, month, int(day_text))
    return candidate.isoformat()


def summarize_records(result: ParseResult) -> str:
    records = list(result.records)
    if not records:
        return "0일치"
    dates = [record.date for record in records]
    origins = {record.origin for record in records if record.origin}
    destinations = {record.destination for record in records if record.destination}
    route = ""
    if len(origins) == 1 and len(destinations) == 1:
        route = f" · {next(iter(origins))}->{next(iter(destinations))}"
    return f"{len(records)}일치 · {min(dates)}~{max(dates)}{route}"
