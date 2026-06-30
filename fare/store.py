"""Firestore REST loader with local cache fallback."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import base64
import json
import ssl
from pathlib import Path
from typing import Any, Mapping
from urllib.error import HTTPError
from urllib.parse import quote
from urllib.request import Request, urlopen


DEFAULT_PROJECT_ID = "fare-calculator-2026"
DEFAULT_API_KEY = "AIzaSyAm-_kZz9Kn4WgeJyVMEyV_TZ3Za2uouFs"


@dataclass(frozen=True)
class FareSnapshot:
    fares: tuple[dict[str, Any], ...]
    seasons: tuple[dict[str, Any], ...]
    loaded_at: str
    source: str
    warning: str | None = None

    @property
    def routes(self) -> tuple[str, ...]:
        return tuple(sorted({str(item.get("route", "")) for item in self.fares if item.get("route")}))


def load_fare_snapshot(
    config: Mapping[str, Any] | None = None,
    cache_path: str | Path | None = None,
    prefer_cache_within_hours: float | None = None,
) -> FareSnapshot:
    config = dict(config or {})
    firestore_config = dict(config.get("firestore") or {})
    project_id = firestore_config.get("project_id") or config.get("firestore_project_id") or DEFAULT_PROJECT_ID
    api_key = firestore_config.get("api_key") or config.get("firestore_api_key") or DEFAULT_API_KEY
    cache = Path(cache_path or config.get("fare_cache_path") or "cache/fares_snapshot.json")

    # 신선한 캐시가 있으면 Firestore 호출을 생략한다 (시작 시 자동 로드의 쿼터 절감용)
    if prefer_cache_within_hours:
        try:
            cached = _load_cache(cache)
        except Exception:
            cached = None
        if cached is not None:
            age = _cache_age_hours(cached)
            if age is not None and 0 <= age <= prefer_cache_within_hours:
                return cached

    try:
        fares, seasons = _fetch_fares_and_seasons(project_id, api_key)
        loaded_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        snapshot = FareSnapshot(fares=fares, seasons=seasons, loaded_at=loaded_at, source="firestore")
        _save_cache(cache, snapshot)
        return snapshot
    except Exception as exc:
        cached = _load_cache(cache)
        if cached is not None:
            return FareSnapshot(
                fares=cached.fares,
                seasons=cached.seasons,
                loaded_at=cached.loaded_at,
                source="cache",
                warning=f"운임 데이터를 새로 불러오지 못해 {cached.loaded_at} 저장본으로 계산합니다. 최신 운임이 반영되지 않았을 수 있습니다. ({exc})",
            )
        raise RuntimeError(f"Firestore 운임 데이터를 불러오지 못했고 캐시도 없습니다: {exc}") from exc


def _fetch_collection(project_id: str, api_key: str, collection: str) -> list[dict[str, Any]]:
    encoded_project = quote(project_id, safe="")
    encoded_collection = quote(collection, safe="/")
    base_url = (
        f"https://firestore.googleapis.com/v1/projects/{encoded_project}"
        f"/databases/(default)/documents/{encoded_collection}"
    )
    items: list[dict[str, Any]] = []
    page_token = ""
    while True:
        url = f"{base_url}?pageSize=300&key={quote(api_key, safe='')}"
        if page_token:
            url += f"&pageToken={quote(page_token, safe='')}"
        payload = _http_json(url)
        items.extend(_decode_document(doc) for doc in payload.get("documents", []))
        page_token = payload.get("nextPageToken") or ""
        if not page_token:
            return items


def _fetch_fares_and_seasons(
    project_id: str, api_key: str
) -> tuple[tuple[dict[str, Any], ...], tuple[dict[str, Any], ...]]:
    """운임/시즌을 단일 스냅샷 문서(meta/*-snapshot)에서 먼저 읽는다(읽기 2회).
    웹앱(fare-calculator)이 저장 시 전체 목록을 이 문서 하나에 함께 발행한다.
    스냅샷이 아직 없으면(404) 기존처럼 컬렉션 전량을 읽는 방식으로 자동 fallback 한다."""
    fares_snapshot = _fetch_snapshot_doc(project_id, api_key, "meta/fares-snapshot")
    seasons_snapshot = _fetch_snapshot_doc(project_id, api_key, "meta/seasons-snapshot")
    if fares_snapshot is not None and seasons_snapshot is not None:
        return tuple(fares_snapshot), tuple(seasons_snapshot)
    fares = tuple(_fetch_collection(project_id, api_key, "fares"))
    seasons = tuple(_fetch_collection(project_id, api_key, "seasons"))
    return fares, seasons


def _fetch_snapshot_doc(project_id: str, api_key: str, doc_path: str) -> list[dict[str, Any]] | None:
    """meta/*-snapshot 단일 문서의 list 필드를 읽어 반환한다.
    문서가 없으면(404) None을 돌려 호출부가 컬렉션 전량 fetch로 fallback 하게 한다.
    429 등 다른 오류는 그대로 전파해 상위에서 캐시 fallback 으로 처리한다."""
    encoded_project = quote(project_id, safe="")
    encoded_path = quote(doc_path, safe="/")
    url = (
        f"https://firestore.googleapis.com/v1/projects/{encoded_project}"
        f"/databases/(default)/documents/{encoded_path}?key={quote(api_key, safe='')}"
    )
    try:
        payload = _http_json(url)
    except HTTPError as exc:
        if exc.code == 404:
            return None
        raise
    item = _decode_document(payload)
    value = item.get("list")
    if isinstance(value, list) and value:
        return [dict(entry) for entry in value if isinstance(entry, dict)]
    return None


_SSL_CONTEXT: ssl.SSLContext | None = None


def _get_ssl_context() -> ssl.SSLContext:
    """Windows 인증서 저장소 기반 컨텍스트를 1회만 만들어 재사용한다.
    requests/certifi를 쓰면 venv가 구글 드라이브에 있을 때 certifi 파일 로드가
    수십 초 걸려 TLS 핸드셰이크가 끊기므로 표준 라이브러리만 사용한다."""
    global _SSL_CONTEXT
    if _SSL_CONTEXT is None:
        _SSL_CONTEXT = ssl.create_default_context()
    return _SSL_CONTEXT


def _http_json(url: str) -> dict[str, Any]:
    req = Request(url, headers={"User-Agent": "NaeilERPUpdaterV5/1.0"})
    try:
        with urlopen(req, timeout=12, context=_get_ssl_context()) as response:
            return json.loads(response.read().decode("utf-8"))
    except HTTPError as exc:
        if exc.code == 429:
            raise RuntimeError(
                "운임 DB 호출 한도 초과(429). Firebase 무료 일일 쿼터가 소진된 상태입니다. "
                "한국시간 오후 4시(태평양 자정)에 초기화되며, 그 전에는 Firebase 콘솔에서 "
                "사용량을 확인해 주세요."
            ) from exc
        raise


def _decode_document(doc: Mapping[str, Any]) -> dict[str, Any]:
    fields = doc.get("fields", {})
    item = {key: _decode_value(value) for key, value in fields.items()}
    name = str(doc.get("name", ""))
    if name:
        item["id"] = name.rsplit("/", 1)[-1]
    return item


def _decode_value(value: Mapping[str, Any]) -> Any:
    if "stringValue" in value:
        return value["stringValue"]
    if "integerValue" in value:
        return int(value["integerValue"])
    if "doubleValue" in value:
        return float(value["doubleValue"])
    if "booleanValue" in value:
        return bool(value["booleanValue"])
    if "timestampValue" in value:
        return value["timestampValue"]
    if "nullValue" in value:
        return None
    if "arrayValue" in value:
        return [_decode_value(item) for item in value.get("arrayValue", {}).get("values", [])]
    if "mapValue" in value:
        return {
            key: _decode_value(inner)
            for key, inner in value.get("mapValue", {}).get("fields", {}).items()
        }
    if "bytesValue" in value:
        return base64.b64decode(value["bytesValue"])
    return None


def _save_cache(path: Path, snapshot: FareSnapshot) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "loadedAt": snapshot.loaded_at,
        "source": snapshot.source,
        "fares": list(snapshot.fares),
        "seasons": list(snapshot.seasons),
    }
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")


def _load_cache(path: Path) -> FareSnapshot | None:
    if not path.exists():
        return None
    payload = json.loads(path.read_text(encoding="utf-8"))
    return FareSnapshot(
        fares=tuple(dict(item) for item in payload.get("fares", [])),
        seasons=tuple(dict(item) for item in payload.get("seasons", [])),
        loaded_at=payload.get("loadedAt") or payload.get("loaded_at") or "",
        source="cache",
    )


def _cache_age_hours(snapshot: FareSnapshot) -> float | None:
    try:
        loaded = datetime.strptime(snapshot.loaded_at, "%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return None
    return (datetime.now() - loaded).total_seconds() / 3600.0
