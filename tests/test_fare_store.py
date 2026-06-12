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
            with mock.patch.object(store, "_fetch_collection", side_effect=AssertionError("네트워크 호출 금지")):
                snap = load_fare_snapshot({}, cache_path=cache, prefer_cache_within_hours=12)
        self.assertEqual(snap.source, "cache")
        self.assertEqual(len(snap.fares), 1)

    def test_stale_cache_fetches_from_network(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            old = (datetime.now() - timedelta(hours=48)).strftime("%Y-%m-%d %H:%M:%S")
            _write_cache(cache, old)
            with mock.patch.object(store, "_fetch_collection", return_value=[{"route": "ICN-SPN"}]):
                snap = load_fare_snapshot({}, cache_path=cache, prefer_cache_within_hours=12)
        self.assertEqual(snap.source, "firestore")

    def test_missing_cache_fetches_from_network(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            with mock.patch.object(store, "_fetch_collection", return_value=[{"route": "ICN-SPN"}]):
                snap = load_fare_snapshot({}, cache_path=cache, prefer_cache_within_hours=12)
        self.assertEqual(snap.source, "firestore")

    def test_force_load_ignores_fresh_cache(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache = Path(tmp) / "snap.json"
            _write_cache(cache, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
            with mock.patch.object(store, "_fetch_collection", return_value=[{"route": "ICN-SPN"}]):
                snap = load_fare_snapshot({}, cache_path=cache)
        self.assertEqual(snap.source, "firestore")


if __name__ == "__main__":
    unittest.main()
