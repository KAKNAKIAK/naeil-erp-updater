import json
import tempfile
import unittest
from datetime import datetime, timedelta
from pathlib import Path
from unittest import mock
from urllib.error import HTTPError

from fare import store
from fare.store import load_fare_snapshot


def _write_cache(path: Path, loaded_at: str) -> None:
    payload = {
        "loadedAt": loaded_at,
        "source": "firestore",
        "fares": [{"route": "ICN-GUM", "price": 100}],
        "seasons": [{"name": "성수기"}],
    }
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")


class HttpJsonErrorTest(unittest.TestCase):
    def test_429_becomes_korean_runtime_error(self):
        err = HTTPError("https://firestore", 429, "Too Many Requests", {}, None)
        with mock.patch.object(store, "urlopen", side_effect=err):
            with self.assertRaises(RuntimeError) as ctx:
                store._http_json("https://firestore")
        self.assertIn("429", str(ctx.exception))
        self.assertIn("한도 초과", str(ctx.exception))

    def test_other_http_errors_propagate(self):
        err = HTTPError("https://firestore", 500, "Server Error", {}, None)
        with mock.patch.object(store, "urlopen", side_effect=err):
            with self.assertRaises(HTTPError):
                store._http_json("https://firestore")


class PreferCacheTest(unittest.TestCase):
    def test_fresh_cache_skips_network(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            _write_cache(cache, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            with mock.patch.object(store, "_fetch_snapshot_doc", side_effect=AssertionError("네트워크 호출 금지")), mock.patch.object(store, "_fetch_collection", side_effect=AssertionError("네트워크 호출 금지")):
                snap = load_fare_snapshot({}, cache_path=cache, prefer_cache_within_hours=12)
        self.assertEqual(snap.source, "cache")
        self.assertEqual(len(snap.fares), 1)

    def test_stale_cache_fetches_from_network(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            old = (datetime.now() - timedelta(hours=48)).strftime("%Y-%m-%d %H:%M:%S")
            _write_cache(cache, old)
            with mock.patch.object(store, "_fetch_snapshot_doc", return_value=None), mock.patch.object(store, "_fetch_collection", return_value=[{"route": "ICN-SPN"}]):
                snap = load_fare_snapshot({}, cache_path=cache, prefer_cache_within_hours=12)
        self.assertEqual(snap.source, "firestore")

    def test_missing_cache_fetches_from_network(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            with mock.patch.object(store, "_fetch_snapshot_doc", return_value=None), mock.patch.object(store, "_fetch_collection", return_value=[{"route": "ICN-SPN"}]):
                snap = load_fare_snapshot({}, cache_path=cache, prefer_cache_within_hours=12)
        self.assertEqual(snap.source, "firestore")

    def test_force_load_ignores_fresh_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            _write_cache(cache, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            with mock.patch.object(store, "_fetch_snapshot_doc", return_value=None), mock.patch.object(store, "_fetch_collection", return_value=[{"route": "ICN-SPN"}]):
                snap = load_fare_snapshot({}, cache_path=cache)
        self.assertEqual(snap.source, "firestore")

    def test_snapshot_doc_used_without_collection_scan(self):
        # 스냅샷 단일 문서가 있으면 컬렉션 전량 스캔(_fetch_collection)을 호출하지 않아야 한다.
        def fake_snapshot(project_id, api_key, doc_path):
            if doc_path == "meta/fares-snapshot":
                return [{"route": "ICN-SPN", "type": "기준", "classCode": "Y", "roundTripFare": 500000}]
            return [{"route": "ICN-SPN", "type": "기준", "startDate": "2026-01-01", "endDate": "2026-12-31"}]

        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            with mock.patch.object(store, "_fetch_snapshot_doc", side_effect=fake_snapshot), mock.patch.object(store, "_fetch_collection", side_effect=AssertionError("스냅샷이 있으면 컬렉션 전량 스캔 금지")):
                snap = load_fare_snapshot({}, cache_path=cache)
        self.assertEqual(snap.source, "firestore")
        self.assertEqual(len(snap.fares), 1)
        self.assertEqual(snap.fares[0]["route"], "ICN-SPN")
        self.assertEqual(len(snap.seasons), 1)


if __name__ == "__main__":
    unittest.main()
