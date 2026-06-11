"""Route inference and forgiving route search for the V5 GUI."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import date
import re
from typing import Iterable

from topas.availability import REQUEST_RE, parse_availability_text


@dataclass(frozen=True)
class RouteInference:
    route: str | None
    reason: str
    candidates: tuple[str, ...] = ()


def normalize_route_query(value: str) -> str:
    return re.sub(r"[^A-Z0-9]+", "", str(value or "").upper())


def filter_routes(routes: Iterable[str], query: str, limit: int | None = 80) -> list[str]:
    route_list = list(routes)
    tokens = _query_tokens(query)
    compact_query = normalize_route_query(query)
    if not tokens and not compact_query:
        return route_list[:limit] if limit is not None else route_list

    scored = []
    for index, route in enumerate(route_list):
        score = _route_match_score(route, tokens, compact_query)
        if score is not None:
            scored.append((score, index, route))

    scored.sort(key=lambda item: (item[0], item[1]))
    matches = [route for _, _, route in scored]
    return matches[:limit] if limit is not None else matches


def _query_tokens(query: str) -> list[str]:
    tokens = []
    for token in re.split(r"[^A-Za-z0-9]+", str(query or "").strip()):
        normalized = normalize_route_query(token)
        if normalized:
            tokens.append(normalized)
    return tokens


def _route_parts(route: str) -> list[str]:
    return [
        normalize_route_query(part)
        for part in re.split(r"[^A-Za-z0-9]+", str(route or "").upper())
        if normalize_route_query(part)
    ]


def _route_match_score(route: str, tokens: list[str], compact_query: str) -> int | None:
    normalized = normalize_route_query(route)
    if not normalized:
        return None

    parts = _route_parts(route)
    score = 0
    if compact_query:
        if normalized == compact_query:
            score -= 50
        elif normalized.startswith(compact_query):
            score -= 8
        elif compact_query in normalized:
            score += 5 + normalized.index(compact_query)
        elif not tokens:
            return None

    for token in tokens:
        token_score = _token_match_score(token, normalized, parts)
        if token_score is None:
            return None
        score += token_score
    return score


def _token_match_score(token: str, normalized_route: str, parts: list[str]) -> int | None:
    scores = []
    for part_index, part in enumerate(parts):
        part_priority = part_index * 3
        if part == token:
            scores.append(part_priority)
        elif part.startswith(token):
            scores.append(10 + part_priority)
        elif token in part:
            scores.append(80 + part_priority + part.index(token))

    if normalized_route.startswith(token):
        scores.append(25)
    elif token in normalized_route:
        scores.append(100 + normalized_route.index(token))

    return min(scores) if scores else None


def infer_route_from_topas_text(
    departure_text: str,
    return_text: str,
    available_routes: Iterable[str],
    year_hint: int | None = None,
) -> RouteInference:
    routes = tuple(available_routes)
    dep = _first_route_info(departure_text, year_hint=year_hint)
    ret = _first_route_info(return_text, year_hint=year_hint)

    candidates: list[tuple[str, str]] = []
    if dep:
        origin, destination, airline = dep
        candidates.append((f"{origin}-{destination}-{airline}", "출발편 원문"))
    if ret and not dep:
        origin, destination, airline = ret
        candidates.append((f"{destination}-{origin}-{airline}", "귀국편 원문 역방향"))
    elif dep and ret:
        dep_origin, dep_destination, dep_airline = dep
        ret_origin, ret_destination, ret_airline = ret
        if dep_airline == ret_airline and dep_origin == ret_destination and dep_destination == ret_origin:
            candidates.insert(0, (f"{dep_origin}-{dep_destination}-{dep_airline}", "왕복 원문"))

    seen = []
    for candidate, reason in candidates:
        if candidate in seen:
            continue
        seen.append(candidate)
        exact = _find_exact_route(routes, candidate)
        if exact:
            return RouteInference(route=exact, reason=reason, candidates=tuple(seen))

        fuzzy = filter_routes(routes, candidate, limit=None)
        if len(fuzzy) == 1:
            return RouteInference(route=fuzzy[0], reason=f"{reason} 유사검색", candidates=tuple(seen))

    return RouteInference(route=None, reason="노선 자동 선택 실패", candidates=tuple(seen))


def _find_exact_route(routes: Iterable[str], candidate: str) -> str | None:
    normalized = normalize_route_query(candidate)
    for route in routes:
        if normalize_route_query(route) == normalized:
            return route
    return None


def _first_route_info(text: str, year_hint: int | None = None) -> tuple[str, str, str] | None:
    if not str(text or "").strip():
        return None
    try:
        blocks = parse_availability_text(text, year_hint=year_hint or date.today().year)
    except Exception:
        return None

    for block in blocks:
        origin = block.origin
        destination = block.destination
        airline = _request_airline(block.request_command)
        if not airline and block.flights:
            airline = block.flights[0].airline
        if origin and destination and airline:
            return origin.upper(), destination.upper(), airline.upper()
    return None


def _request_airline(request_command: str | None) -> str | None:
    match = REQUEST_RE.search(str(request_command or "").strip())
    if not match:
        return None
    airline = match.group("airline")
    return airline.upper() if airline else None
