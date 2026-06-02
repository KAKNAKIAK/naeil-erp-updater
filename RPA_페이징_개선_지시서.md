# ERP Updater RPA 페이징 처리 개선 지시서 (v3.0.1)

> 개정 이력
> - **v3.0.1 (2026-06-02)**: 실제 디버깅 크롬(포트 9222) 세션에서 ERP 그리드를 직접 계측해 초안의 가정을 검증·수정. `next` 버튼 누적 클릭 → `moveToPage` 직접 점프로 전략 변경, `totalPage` 지연 버그 차단(totRows 기반 계산), 페이지 이동 검증 추가. `gui.py`에 구현 완료.
> - 초안: next 버튼 반복 클릭 + `totalPage` 변수 기반 종료 판정(미검증).

본 문서는 내일투어 ERP(erp.naeiltour.co.kr) **행사요금일괄수정** 화면에서 기간 조회 결과가 500건을 초과할 때, 모든 페이지를 누락 없이 순회하며 요금을 저장하기 위한 RPA 명세서입니다. 모든 셀렉터·변수·함수 동작은 실제 브라우저 계측으로 확인된 값입니다.

---

## 1. 배경 및 문제 정의

- **현상**: 행사요금일괄수정 화면에서 넓은 날짜 범위를 조회하면 한 페이지에 **정확히 500건**만 로드되고, 하단에 페이징 버튼이 활성화됩니다.
- **기존 한계**: RPA가 1페이지 500건만 저장하고 종료 → 501번째 이후 행사가 누락됩니다.
- **목표(제1원칙)**: **누락 0건.** 조회 결과가 여러 페이지면 자동으로 끝 페이지까지 전부 저장합니다. 중간에 페이지 이동이 어긋나면 "조용히 건너뛰기" 대신 **명시적 실패 처리**로 중단해, 누락을 숨기지 않습니다.

---

## 2. 실측 검증 결과 (포트 9222 직접 계측)

테스트 조건: 출발일 `2025-10-01 ~ 2025-10-31` 조회 → 총 **3,536건 / 8페이지**.

### 2.1 그리드·iframe 구조
- 대상 iframe: `#framego01_104_list` (`…/erp/go/go01/pop/go01_104_list`). 메인과 같은 도메인이라 제어 가능.
- AUI Grid id: **`#gridMain`**, 페이지당 **500행** 고정.
- 행 데이터에 **`totRows`**(전체 건수, 예: 3536)가 모든 행에 즉시 포함됨. 고유키 후보: `basePriceSeq`(행 PK), `eventSeq`, `eventCd`, `goodSeq`.

### 2.2 페이징 버튼
- 다음: `#paging-button-next` / 이전: `#paging-button-prev`
  - 클래스 오타까지 실제와 일치: `aui-grid-paging-simple-button aui-grid-paging-simple-butotn-next`
  - 데이터 없음·단일 페이지일 때 `style="display:none"`로 숨김. 노출 여부는 `offsetParent !== null`로 판정.

### 2.3 페이지 이동 함수 (실제 원문)
```javascript
// 다음 버튼 onclick:
//   moveToPage(Math.min(currentPage + 1, totalPage), 'this', '#gridMain_r');

function moveToPage(goPage, parentFg, rButton) {
    createPagingNavigator(goPage);          // prev/next 버튼 표시 토글
    currentPage = goPage;                   // 현재 페이지를 goPage로 직접 세팅
    pageClick = true;                       // 저장 후 1페이지 리셋을 막는 플래그
    if (parentFg == "parent") { parent.$(rButton).click(); }
    else { $(rButton).click(); }            // 'this' 등 → iframe 내부에서 조회 재실행
}
```
- **서버 사이드 페이징**: `moveToPage`는 `currentPage`를 바꾼 뒤 조회 버튼(`#gridMain_r`)을 다시 눌러 서버에서 해당 페이지를 받아옵니다.
- **핵심 활용**: `moveToPage(N, 'this', '#gridMain_r')` 한 줄로 **현재 페이지와 무관하게 N페이지로 직접 점프**합니다. (1→5 점프 실측 확인.) 즉 next 버튼을 (N−1)번 누적 클릭할 필요가 없습니다.

### 2.4 저장 후 1페이지 리셋
- 일반 조회(저장 후 자동 재조회 포함)는 `pageClick=false` 상태라 `currentPage`가 **1로 리셋**됩니다.
- 단, `moveToPage`는 `pageClick=true`로 호출하므로 이동 후 `currentPage`가 목표값에 유지됩니다(리셋 안 됨). → 이동 성공 검증에 `currentPage == target_page`를 그대로 쓸 수 있음.

### 2.5 ⚠️ 초안 대비 정정 사항 (누락 직결)
1. **`totalPage` 변수는 행 바인딩보다 늦게 채워진다.** 조회 직후 `rowCount=500`이어도 `totalPage`는 약 5~8초간 `0`이다가 뒤늦게 `8`로 갱신됨.
   - 초안처럼 grid-ready 직후 `totalPage`를 읽어 종료 판정하면 `0 → (or 1) → 즉시 break`로 **2~8페이지가 통째로 누락**됨. **반드시 `totRows`로 전체 페이지를 계산할 것.**
2. **`usePaging`, `gridData` 전역은 존재하지 않음(undefined).** 초안의 "`gridData.page=1`로 리셋" 표현은 부정확. 리셋 동작 자체는 `pageClick` 플래그로 동작.

---

## 3. 제어 흐름 (구현된 알고리즘)

```
[날짜/기간 루프]
 └ 1. 시작일/종료일 주입 → 조회 클릭
 └ 2. 날짜가 검색범위에 들어오는지(matched) 폴링 = 1페이지 바인딩 확인
 └ 3. totRows 읽어 total_pages = ceil(totRows / 500) 계산   ← totalPage 변수 미사용
 └ 4. [페이지 루프] target_page = 1 .. total_pages
        ├ 4-a. target_page > 1 이면 navigate_to_grid_page(target_page)
        │        · moveToPage(target_page,'this','#gridMain_r') 직접 점프
        │        · 그리드 락 해제 대기 → currentPage == target_page 검증(최대 3회)
        │        · 검증 실패 시 RuntimeError로 이 기간 처리 중단(누락 은폐 금지)
        ├ 4-b. 전체선택 → 요금업데이트 모달 → 값 주입 → 저장 전 검증 → 저장 → 얼럿 수락 → 모달 닫기
        ├ 4-c. wait_until_grid_ready_after_save (저장 후 1페이지 리셋 완료 대기)
        └ 4-d. target_page += 1
 └ 5. 전체 페이지 처리 완료 → SUCCESS 기록
```

핵심 차이: 저장하면 1페이지로 튕기지만, 다음 페이지는 **현재 위치와 무관한 직접 점프**라 클릭 누적·락 어긋남으로 인한 페이지 오인이 발생하지 않습니다.

---

## 4. 구현 내역 (`gui.py`, v3.0.1)

### 4.1 config.json 추가 키
```json
"grid_id": "#gridMain",
"grid_page_size": 500,
"paging_search_button": "#gridMain_r"
```

### 4.2 `get_grid_page_state()` — 페이징 상태 획득
`totalPage` 변수 대신 `AUIGrid.getGridData('#gridMain')[0].totRows`로 전체 건수를 읽어 페이지 수를 계산합니다. (totRows는 행과 동시에 채워져 지연 없음.)

```python
def get_grid_page_state(self):
    page_size = int(self.config.get('grid_page_size', 500)) or 500
    grid_id = self.config.get('grid_id', '#gridMain')
    js = """
    try {
        var gid = arguments[0];
        var d = (typeof AUIGrid !== 'undefined') ? AUIGrid.getGridData(gid) : null;
        var tot = (d && d.length) ? (d[0].totRows || d.length) : 0;
        var nextBtn = document.querySelector('#paging-button-next');
        return {
            cur: (typeof currentPage !== 'undefined' ? currentPage : 1),
            totRows: tot, pageRows: (d ? d.length : 0),
            nextVisible: !!(nextBtn && nextBtn.offsetParent !== null)
        };
    } catch (e) { return {cur:1, totRows:0, pageRows:0, nextVisible:false}; }
    """
    st = self.driver.execute_script(js, grid_id) or {}
    tot, cur = int(st.get('totRows') or 0), int(st.get('cur') or 1)
    page_rows, next_visible = int(st.get('pageRows') or 0), bool(st.get('nextVisible'))
    if tot > 0:
        total_pages = max(1, -(-tot // page_size))           # ceil
    elif page_rows > 0 and next_visible:
        total_pages = max(2, cur + 1)                        # totRows 실패 시 안전측
    else:
        total_pages = 1
    return cur, total_pages, tot, next_visible
```

### 4.3 `navigate_to_grid_page()` — 목표 페이지 직접 점프 + 검증
```python
def navigate_to_grid_page(self, selectors, target_page, timeout):
    rbutton = self.config.get('paging_search_button', '#gridMain_r')
    cur = -1
    for attempt in range(3):
        if not self.is_running:
            return False
        self.driver.switch_to.default_content()
        self.find_and_switch_frame(selectors["search_date_input"])
        self.driver.execute_script(
            "try { moveToPage(arguments[0], 'this', arguments[1]); } catch (e) {}",
            target_page, rbutton)
        self.wait_until_grid_ready_after_save(selectors, timeout)
        cur, _, _, _ = self.get_grid_page_state()
        if cur == target_page:
            return True
        print(f" -> [페이지 이동 재시도] 목표 {target_page} / 현재 {cur} (시도 {attempt+1}/3)")
        time.sleep(0.5)
    raise RuntimeError(f"{target_page}페이지로 이동하지 못했습니다(현재 {cur}). "
                       f"데이터 누락을 막기 위해 이 기간 처리를 중단합니다.")
```

### 4.4 `rpa_worker_loop` 통합
기존 "전체선택 → 모달 → 입력 → 검증 → 저장 → 모달 닫기 → grid-ready 대기" 블록을 **그대로 보존**한 채 페이지 `while` 루프로 감쌌습니다. 저장 로직(값 주입·이중 검증)은 변경하지 않았습니다. 2페이지 이후만 루프 진입 시 `navigate_to_grid_page`로 이동합니다.

---

## 5. 검증·예외 처리 원칙

1. **단일 페이지(≤500건) 회귀**: `total_pages == 1`이면 루프 1회 후 종료 → 기존 v3.0.0 동작과 동일. (회귀 테스트 필수)
2. **totalPage 지연 무관**: totRows 기반이라 totalPage가 0이든 늦든 영향 없음.
3. **페이지 이동 실패 = 명시적 실패**: `currentPage` 검증 3회 실패 시 해당 기간을 FAIL 처리하고 중단. 누락을 SUCCESS로 위장하지 않음.
4. **사용자 중지 대응**: 페이지 루프 중 `is_running=False`면 즉시 멈추고, 부분 처리 상태를 실패로 기록.
5. **락 대기 생략 금지**: 매 페이지 이동·저장 후 `wait_until_grid_ready_after_save`(스마트 폴링)로 그리드 마스크 해제를 확인.

### 5.1 멤버십 변동 검증 — 완료 (2026-06-02)
- **검증 방법**: 무손상 단일행 동일값 저장. `2025-10-01~31`(3,536건/8페이지) 1페이지 첫 행(`basePriceSeq=34635700`)만 체크 → 현재 요금을 그대로 다시 입력(성분합=adultPrice 게이트 통과) → 저장 커밋 → 동일 기간 재조회.
- **결과**: `totRows` 3536 → 3536(불변), 저장행이 **결과셋에 그대로 잔존(같은 위치 index 0)**, `adultPrice` 642192 변동 없음.
- **결론**: 저장해도 행이 검색결과에서 빠지지 않음 → **인덱스 기반 페이지 순회 안전**. 페이지가 밀려 누락될 위험 없음.
- (참고) 만약 향후 ERP가 "수정완료 제외" 같은 필터를 도입해 멤버십이 변하게 되면, `basePriceSeq` 처리키 추적(처리 키 Set 기록 → 미처리 키만 저장)으로 전환하면 누락·중복 0 보장.

---

## 6. 테스트 체크리스트 (v3.0.1)

CDP 직접 계측 + 무손상 검증으로 아래 확인 완료(2026-06-02):
- [x] 멀티 페이지(3,536건/8페이지): `ceil(totRows/500)=8` 정확, 8페이지 전부 `moveToPage` 점프·전체선택·모달 진입 성공(행수 500×7 + 36)
- [x] `moveToPage` 점프 간헐 실패(5페이지→1페이지) 재현 → `currentPage` 검증+3회 재시도로 복구(5/5) → **검증 로직이 누락을 실제로 차단**
- [x] 멤버십 변동 없음(5.1) → 인덱스 순회 안전
- [x] 무손상 단일행 동일값 저장: adultPrice·totRows 불변 확인
- [ ] (실사용 권장) 실제 엑셀로 GUI 멀티페이지 1회 구동 — 콘솔에 `[페이지 k/N]` 로그가 N개 찍히는지 최종 육안 확인
- [ ] (회귀) 단일 페이지(≤500건) 기간: `[페이징]`/`[페이지 이동]` 로그 없이 1회 저장 후 종료
