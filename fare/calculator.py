"""Air fare calculation compatible with fare-calculator-2026."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timedelta
import math
from typing import Iterable, Mapping, Sequence

from .parser import ParsedAvailabilityRecord, parse_topas_text


@dataclass(frozen=True)
class OneWayDebug:
    date: str
    classes: tuple[str, ...]
    class_fares: tuple[dict[str, object], ...]
    season: dict[str, object] | None
    season_type: str
    selected_class: str
    final_fare: float
    is_closed: bool


@dataclass(frozen=True)
class RoundTripResult:
    dep_date: str
    ret_date: str
    nights: int
    total_fare: float
    is_closed: bool
    dep_fare: float
    ret_fare: float


@dataclass(frozen=True)
class CombinationDebug:
    dep_date: str
    ret_date: str
    nights: int
    dep_fare: float
    ret_fare: float
    total_fare: float
    is_closed: bool
    dep_season: str
    ret_season: str


@dataclass(frozen=True)
class CalculationDebug:
    dep_records: tuple[ParsedAvailabilityRecord, ...]
    ret_records: tuple[ParsedAvailabilityRecord, ...]
    dep_debug: tuple[OneWayDebug, ...]
    ret_debug: tuple[OneWayDebug, ...]
    combinations: tuple[CombinationDebug, ...]
    fare_map: dict[str, dict[str, float]]
    seasons: tuple[dict[str, object], ...]


@dataclass(frozen=True)
class CalculationResult:
    result: dict[int, list[RoundTripResult]]
    debug: CalculationDebug
    warnings: tuple[str, ...] = field(default_factory=tuple)


def js_round(value: float) -> int:
    return int(math.floor(float(value) + 0.5))


def calculate_round_trips(
    dep_text: str,
    ret_text: str,
    fares: Iterable[Mapping[str, object]],
    seasons: Iterable[Mapping[str, object]],
    route: str,
    nights: Sequence[int] = (2, 3, 4),
    today=None,
) -> CalculationResult:
    dep_parse = parse_topas_text(dep_text, today=today)
    ret_parse = parse_topas_text(ret_text, today=today)
    warnings: list[str] = [*dep_parse.warnings, *ret_parse.warnings]

    fare_map, fare_warnings = build_fare_map(fares, route)
    warnings.extend(fare_warnings)
    route_seasons = tuple(
        dict(item) for item in seasons if str(item.get("route", "")).strip() == route
    )

    dep_calc = _build_one_way_maps(dep_parse.records, route_seasons, fare_map, warnings)
    ret_calc = _build_one_way_maps(ret_parse.records, route_seasons, fare_map, warnings)

    ret_fare_map = ret_calc["fare_map"]
    ret_closed_map = ret_calc["closed_map"]
    dep_closed_map = dep_calc["closed_map"]
    dep_debug_by_date = {item.date: item for item in dep_calc["debug"]}
    ret_debug_by_date = {item.date: item for item in ret_calc["debug"]}

    result: dict[int, list[RoundTripResult]] = {int(night): [] for night in nights}
    combinations: list[CombinationDebug] = []
    out_of_range: dict[int, list[str]] = {int(night): [] for night in nights}

    for dep_date in sorted(dep_calc["fare_map"].keys()):
        dep_fare = dep_calc["fare_map"][dep_date]
        if not dep_fare:
            continue
        for night in result.keys():
            ret_date = add_days(dep_date, night)
            ret_fare = ret_fare_map.get(ret_date)
            if not ret_fare:
                out_of_range[night].append(dep_date)
                continue
            closed = bool(dep_closed_map.get(dep_date) or ret_closed_map.get(ret_date))
            total = dep_fare + ret_fare
            row = RoundTripResult(
                dep_date=dep_date,
                ret_date=ret_date,
                nights=night,
                total_fare=total,
                is_closed=closed,
                dep_fare=dep_fare,
                ret_fare=ret_fare,
            )
            result[night].append(row)
            combinations.append(
                CombinationDebug(
                    dep_date=dep_date,
                    ret_date=ret_date,
                    nights=night,
                    dep_fare=dep_fare,
                    ret_fare=ret_fare,
                    total_fare=total,
                    is_closed=closed,
                    dep_season=dep_debug_by_date.get(dep_date).season_type
                    if dep_date in dep_debug_by_date
                    else "",
                    ret_season=ret_debug_by_date.get(ret_date).season_type
                    if ret_date in ret_debug_by_date
                    else "",
                )
            )

    for night, dates in out_of_range.items():
        if dates:
            warnings.append(
                f"{night}박 기준 {len(dates)}건은 귀국편 원문 범위를 벗어나 제외했습니다 "
                f"({min(dates)}~{max(dates)} 출발)."
            )

    return CalculationResult(
        result=result,
        debug=CalculationDebug(
            dep_records=dep_parse.records,
            ret_records=ret_parse.records,
            dep_debug=tuple(dep_calc["debug"]),
            ret_debug=tuple(ret_calc["debug"]),
            combinations=tuple(combinations),
            fare_map=fare_map,
            seasons=route_seasons,
        ),
        warnings=tuple(warnings),
    )


def build_fare_map(
    fares: Iterable[Mapping[str, object]],
    route: str,
) -> tuple[dict[str, dict[str, float]], list[str]]:
    fare_map: dict[str, dict[str, float]] = {}
    seen: set[tuple[str, str]] = set()
    duplicates: set[tuple[str, str]] = set()

    for item in fares:
        if str(item.get("route", "")).strip() != route:
            continue
        fare_type = str(item.get("type", "기준")).strip() or "기준"
        class_code = str(item.get("classCode", "")).strip().upper()
        if not class_code:
            continue
        round_trip = _to_number(item.get("roundTripFare"))
        if round_trip is None:
            continue
        key = (fare_type, class_code)
        if key in seen:
            duplicates.add(key)
        seen.add(key)
        fare_map.setdefault(fare_type, {})[class_code] = round_trip / 2

    warnings = []
    if duplicates:
        warnings.append(
            "중복 운임 문서는 마지막 값을 사용했습니다: "
            + ", ".join(f"{fare_type}/{code}" for fare_type, code in sorted(duplicates))
        )
    return fare_map, warnings


def _build_one_way_maps(
    records: Iterable[ParsedAvailabilityRecord],
    seasons: Sequence[Mapping[str, object]],
    fare_map: dict[str, dict[str, float]],
    warnings: list[str],
) -> dict[str, object]:
    base_map = fare_map.get("기준", {})
    fare_by_date: dict[str, float] = {}
    closed_by_date: dict[str, bool] = {}
    debug_rows: list[OneWayDebug] = []

    for record in records:
        season = find_season(record.date, seasons, warnings)
        season_type = str(season.get("type")) if season else "기준"
        season_map = fare_map.get(season_type, {})
        class_fares = []
        for token in record.classes:
            code = reclass(token)
            class_fares.append(
                {
                    "cls": token,
                    "code": code,
                    "fare": season_map.get(code, base_map.get(code)),
                    "seasonFare": season_map.get(code),
                    "baseFare": base_map.get(code),
                }
            )
        selected_class, fare, is_closed = select_lowest_registered_class(
            record.classes,
            season_type,
            fare_map,
        )
        fare_by_date[record.date] = fare
        closed_by_date[record.date] = is_closed
        debug_rows.append(
            OneWayDebug(
                date=record.date,
                classes=record.classes,
                class_fares=tuple(class_fares),
                season=dict(season) if season else None,
                season_type=season_type,
                selected_class=selected_class,
                final_fare=fare,
                is_closed=is_closed,
            )
        )

    return {
        "fare_map": fare_by_date,
        "closed_map": closed_by_date,
        "debug": tuple(debug_rows),
    }


def find_season(
    travel_date: str,
    seasons: Sequence[Mapping[str, object]],
    warnings: list[str] | None = None,
) -> Mapping[str, object] | None:
    matches = [
        item
        for item in seasons
        if str(item.get("startDate", "")) <= travel_date <= str(item.get("endDate", ""))
    ]
    if len(matches) > 1 and warnings is not None:
        warnings.append(
            f"{travel_date} 시즌 구간이 {len(matches)}개 겹쳐 첫 구간({matches[0].get('type')})을 사용했습니다."
        )
    return matches[0] if matches else None


def select_lowest_registered_class(
    classes: Iterable[str],
    season_type: str,
    fare_map: dict[str, dict[str, float]],
) -> tuple[str, float, bool]:
    base_map = fare_map.get("기준", {})
    season_map = fare_map.get(season_type, {})
    base_y = base_map.get("Y")
    fallback = base_y * 1.2 if base_y else 0
    lowest = float("inf")
    selected = ""
    for token in classes:
        code = reclass(token)
        fare = season_map.get(code, base_map.get(code))
        if fare is not None and fare < lowest:
            lowest = fare
            selected = code
    if lowest == float("inf"):
        return "N/A", fallback, True
    return selected, lowest, False


def reclass(token: str) -> str:
    return "".join(ch for ch in str(token).upper() if not ch.isdigit())


def add_days(iso_date: str, days: int) -> str:
    dt = datetime.strptime(iso_date, "%Y-%m-%d").date()
    return (dt + timedelta(days=int(days))).isoformat()


def _to_number(value) -> float | None:
    try:
        text = str(value).replace(",", "").strip()
        if not text:
            return None
        return float(text)
    except (TypeError, ValueError):
        return None
