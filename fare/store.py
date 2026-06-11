"""Firestore REST loader with local cache fallback."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime
import base64
import json
from pathlib import Path
from typing import Any, Mapping
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


def load_fare_snapshot(config: Mapping[str, Any] | None = None, cache_path: str | Path | None = None) -> FareSnapshot:
    config = dict(config or {})
    firestore_config = dict(config.get("firestore") or {})
    project_id = firestore_config.get("project_id") or config.get("firestore_project_id") or DEFAULT_PROJECT_ID
    api_key = firestore_config.get("api_key") or config.get("firestore_api_key") or DEFAULT_API_KEY
    cache = Path(cache_path or config.get("fare_cache_path") or "cache/fares_snapshot.json")

    try:
        fares = tuple(_fetch_collection(project_id, api_key, "fares"))
        seasons = tuple(_fetch_collection(project_id, api_key, "seasons"))
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


def _http_json(url: str) -> dict[str, Any]:
    try:
        import requests

        response = requests.get(url, timeout=12)
        response.raise_for_status()
        return response.json()
    except ImportError:
        req = Request(url, headers={"User-Agent": "NaeilERPUpdaterV5/1.0"})
        with urlopen(req, timeout=12) as response:
            return json.loads(response.read().decode("utf-8"))


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
