# 운임/시즌 DB 로드 실패 — 원인 진단 및 수정 지시서

> 작성: 2026-06-12 / 대상: `erp_updater` (NaeilERPUpdater V5)
> 증상: 토파스 요금조회 탭에서 `운임/시즌 로드 실패` 표시, 요금 계산 불가
>
> **반영 상태 (2026-06-12)**: [필수 1]·[필수 2]·[선택 4] 코드 반영 완료. 전체 테스트 31개 통과. v5.0.1 설치본 빌드 및 silent 설치 검증 완료.
> 실서버 확인 결과 0.3초 만에 Firestore 응답 수신(TLS 문제 해소), 현재는 429 안내 메시지가 정상 출력됨.
> 남은 일: Firebase 콘솔 사용량 점검.

## 한줄 결론

원인은 두 가지가 겹쳤다. ① `fare/store.py`가 `requests`로 Firestore를 호출할 때 구글 드라이브(G:) 안의 certifi 인증서 파일을 읽다가 TLS 연결이 끊기고(이 PC에서 100% 재현), ② Firestore 프로젝트 자체가 현재 호출 한도 초과(429 Quota exceeded) 상태다. 캐시 폴백도 한 번도 성공한 적이 없어 `cache/fares_snapshot.json`이 존재하지 않아 최종 실패로 떨어진다.

## 진단 과정에서 확인된 사실 (2026-06-12 오전, 이 PC 기준)

| 테스트 | 결과 |
|---|---|
| `requests.get(firestore)` | **실패** — `SSLEOFError: UNEXPECTED_EOF_WHILE_READING` (핸드셰이크 중 서버가 연결 종료) |
| `requests.get(google.com / github.com)` | 동일하게 실패 → Firestore 문제가 아니라 requests 자체 문제 |
| 표준 라이브러리 `urllib.request.urlopen(firestore)` | **연결 성공** — 서버 응답 HTTP 429 수신 |
| `urllib3.PoolManager()` 직접 호출 | 연결 성공 — HTTP 429 |
| `urllib3.PoolManager(ca_certs=certifi.where())` | requests와 동일하게 **실패** |
| `requests.get(..., verify=False)` | **성공** — HTTP 429 |
| `ssl_context.load_verify_locations(certifi.where())` 단독 실행 | **27.7초 소요** (파일 자체 읽기는 0.01초) |
| `curl.exe` (Windows schannel) | 연결 성공 — HTTP 429 `"Quota exceeded." RESOURCE_EXHAUSTED` (수 분 간격 재시도에도 동일) |
| 프록시 설정 | 없음 (env/winhttp/레지스트리 모두 직결) |
| `cache/fares_snapshot.json` | **폴더째 없음** → 캐시 폴백 불가 |

certifi 경로: `G:\내 드라이브\안티그래비티\erp_updater\.venv\Lib\site-packages\certifi\cacert.pem`

## 원인 정리

### 원인 1 — requests + 구글 드라이브 위 certifi 조합의 TLS 실패 (클라이언트, 이 PC)

`fare/store.py`의 `_http_json()`은 requests를 우선 사용한다. requests는 인증서 검증에 venv 안의 `certifi/cacert.pem`을 쓰는데, venv가 구글 드라이브(G:)에 있어 OpenSSL의 `load_verify_locations()` 호출이 약 28초 걸린다. urllib3는 TCP 연결을 먼저 열고 나서 인증서를 로드하므로, 로드가 끝나기 전에 구글 서버가 유휴 연결을 끊어 `UNEXPECTED_EOF_WHILE_READING`으로 실패한다. 표준 라이브러리 `urllib`은 Windows 인증서 저장소를 쓰기 때문에 같은 PC에서 정상 동작한다(같은 이유로 `update_client.py`의 업데이트 확인은 멀쩡하다).

현재 코드는 requests가 "설치돼 있으면" 무조건 쓰고, `ImportError`일 때만 urllib으로 폴백하므로 이 PC에서는 항상 실패 경로를 탄다.

```python
# fare/store.py 현재 코드 (79~89행)
def _http_json(url: str) -> dict[str, Any]:
    try:
        import requests
        response = requests.get(url, timeout=12)   # ← 여기서 SSLEOFError
        response.raise_for_status()
        return response.json()
    except ImportError:                            # ← SSLError는 안 잡힘 → 폴백 안 됨
        ...
```

### 원인 2 — Firestore 쿼터 초과 429 (서버, 모든 PC 공통)

TLS를 우회해서 접속해 봐도 Firestore가 `429 RESOURCE_EXHAUSTED ("Quota exceeded.")`를 반환한다. `fare-calculator-2026` 프로젝트의 무료(Spark) 일일 읽기 한도가 소진된 것으로 추정된다. 즉 지금은 원인 1을 고쳐도 당장은 로드가 안 된다.

쿼터 소모 구조도 문제다. GUI는 시작할 때마다(`gui.py` 1423행, 시작 0.7초 후 자동) `fares`·`seasons` **전체 컬렉션**을 읽고, ↻ 버튼을 누를 때마다 또 전체를 읽는다. 운임 계산 사이트(fare-calculator-2026.web.app)도 같은 DB를 쓰므로 직원 수 × 실행 횟수만큼 읽기가 누적된다.

무료 쿼터는 태평양 시간 자정(한국시간 **오후 4시**, 서머타임 기준)에 초기화되므로 오늘 오후 4시 이후 자동 복구될 가능성이 높다.

### 원인 3 — 캐시 폴백 부재 (보조)

캐시 파일은 "성공한 로드" 후에만 생성되는데, 이 환경에서는 한 번도 성공한 적이 없어 폴백할 캐시가 없다. 그래서 위 두 원인이 곧바로 `운임/시즌 로드 실패`로 이어진다.

## 수정 지시

### [필수 1] `fare/store.py` — requests 제거, 표준 라이브러리로 통일

`_http_json()`을 아래로 교체한다. 핵심: ① certifi 의존(requests) 제거, ② SSL 컨텍스트를 모듈에서 1회만 만들어 재사용, ③ 429를 사용자가 알아볼 수 있는 한국어 메시지로 변환.

```python
# 상단 import에 추가
import ssl
from urllib.error import HTTPError

_SSL_CONTEXT: ssl.SSLContext | None = None


def _get_ssl_context() -> ssl.SSLContext:
    """Windows 인증서 저장소 기반 컨텍스트를 1회만 만들어 재사용한다.
    (venv가 구글 드라이브에 있으면 certifi 파일 로드가 수십 초 걸려
    TLS 핸드셰이크가 끊기므로 requests/certifi를 쓰지 않는다)"""
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
```

`try: import requests` 분기는 통째로 삭제한다. 이 모듈 외에 requests를 쓰는 곳은 없음을 확인했다(`update_client.py`는 이미 urllib 사용).

### [필수 2] 시작 시 자동 로드를 캐시 우선으로 — 쿼터 소모 절감

매 실행마다 전체 컬렉션을 읽는 구조가 429를 유발/악화시킨다. 캐시가 신선하면 Firestore를 건너뛰게 한다.

`fare/store.py`의 `load_fare_snapshot()`에 파라미터 추가:

```python
def load_fare_snapshot(config=None, cache_path=None, prefer_cache_within_hours=None):
    ...
    cache = Path(cache_path or config.get("fare_cache_path") or "cache/fares_snapshot.json")

    # 신선한 캐시가 있으면 네트워크 호출 생략 (시작 시 자동 로드용)
    if prefer_cache_within_hours:
        cached = _load_cache(cache)
        if cached is not None and _cache_age_hours(cached) is not None \
                and _cache_age_hours(cached) <= prefer_cache_within_hours:
            return cached

    try:
        fares = ...
```

보조 함수:

```python
def _cache_age_hours(snapshot: FareSnapshot) -> float | None:
    try:
        loaded = datetime.strptime(snapshot.loaded_at, "%Y-%m-%d %H:%M:%S")
    except (TypeError, ValueError):
        return None
    return (datetime.now() - loaded).total_seconds() / 3600.0
```

`gui.py` 연결 (3곳):

```python
# 1423행: 시작 자동 로드 → 캐시 우선
self.root.after(700, lambda: self.refresh_fare_snapshot(force=False))

# 1838행: ↻ 버튼은 기본값 force=True → 항상 새로 받음
def refresh_fare_snapshot(self, force=True):
    ...
    thread = threading.Thread(target=self._fare_snapshot_worker, args=(force,), daemon=True)

# 1846행: 워커에서 전달
def _fare_snapshot_worker(self, force=True):
    ...
    snapshot = load_fare_snapshot(
        self.config,
        cache_path=cache_path,
        prefer_cache_within_hours=None if force else 12,
    )
```

기준 12시간은 운임 갱신 주기에 맞춰 조정 가능. 캐시 사용 시 상태표시는 기존 로직이 이미 `운임/시즌 캐시 사용`으로 표기하므로 추가 작업 없음.

### [선택 3] 429 재시도 백오프

분당 한도 같은 일시적 429 대비로 `_fetch_collection()` 호출 전후에 1~2회(2초, 5초) 재시도를 넣을 수 있다. 단, 일일 쿼터 소진이면 재시도해도 소용없으므로 메시지 안내([필수 1])가 우선이다.

### [선택 4] 테스트 추가

`tests/test_fare_store.py` 신설: ① `_http_json`이 429 `HTTPError`를 한국어 `RuntimeError`로 변환하는지, ② `prefer_cache_within_hours` 적용 시 신선 캐시가 그대로 반환되는지 mock으로 검증.

## 코드 외 운영 조치

1. **Firebase 콘솔 확인(권장, 오늘)**: [console.firebase.google.com](https://console.firebase.google.com/project/fare-calculator-2026/usage) → Firestore 사용량에서 읽기 횟수 추이를 확인한다. 어떤 날 5만 회(무료 한도)를 넘었는지, 웹앱/GUI 어느 쪽 트래픽인지 파악. 반복되면 Blaze 전환 또는 읽기 절감([필수 2])으로 대응.
2. **오늘의 임시 복구**: 한국시간 오후 4시 쿼터 리셋 후 ↻ 버튼으로 재로드. [필수 1]이 반영돼 있어야 이 PC에서 통신이 된다.
3. **venv 위치(참고)**: [필수 1]로 certifi 의존이 없어지면 급하지 않지만, 구글 드라이브 위 venv는 같은 류의 문제를 또 만들 수 있다(naeil-erp-mcp를 C:\mcp로 옮긴 것과 같은 이유). 여유 있을 때 C: 이전 권장.
4. **배포본 영향**: PyInstaller 패키징본은 certifi가 C: 임시폴더에 풀려 원인 1의 영향은 없지만, 원인 2(429)는 모든 배포 PC에 공통이다. [필수 1]·[필수 2] 반영 후 v5.0.1로 재배포 권장.

## 수정 후 검증 방법

```powershell
# erp_updater 폴더에서
.venv\Scripts\python.exe -m py_compile gui.py fare\store.py
.venv\Scripts\python.exe -m unittest discover -s tests -v

# 실동작: 쿼터 리셋(오후 4시) 후 GUI 실행 →
#  1) 토파스 요금조회 탭 상태가 '운임/시즌 준비됨'(녹색)인지
#  2) cache\fares_snapshot.json 생성되는지
#  3) GUI 재시작 시 12시간 이내면 '운임/시즌 캐시 사용'으로 즉시 뜨는지
#  4) ↻ 버튼은 항상 Firestore에서 새로 받는지
# 429 상태에서 실행하면 로그에 '운임 DB 호출 한도 초과(429)...' 한국어 안내가 떠야 함
```
