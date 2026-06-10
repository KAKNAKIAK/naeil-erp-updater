# v4.0 TOPAS 좌석조회 기능 작업 메모

## 목적

기존 v3.x ERP 항공요금 수정 RPA는 그대로 유지하고, v4.0에서 TOPAS SellConnect 좌석조회 결과 수집 기능을 별도 탭으로 추가한다.

이번 기능의 핵심은 사용자가 첫 날짜 조회를 직접 완료한 뒤, 프로그램이 **AC1을 10건씩 묶어 전송 -> 묶음 전체 응답 완료 감지 -> 다음 묶음 전송 -> 전체 원문 결과 팝업 표시** 흐름으로 처리하는 것이다.

## 원프로그램 보호 원칙

- 원본 작업트리 `G:\내 드라이브\안티그래비티\erp_updater`는 v3.0.3 안정판으로 유지한 채 별도 worktree에서 v4 기능을 검증했다.
- v4.0 정식 반영 대상은 `erp_updater_v4_topas_lab`의 `feature/v4-topas-seat-checker` 브랜치에서 준비했다.
- 기존 ERP 저장 루프는 그대로 유지하고, TOPAS 기능은 별도 `토파스조회` 탭과 `topas/` 패키지로 분리했다.

## 관찰된 업무 흐름

1. SellConnect Entry 화면에서 최초 `AN...ICNGUM/ALJ915` 좌석조회를 실행한다.
2. 매크로가 `AC1` 또는 다음 날짜 조회 명령을 빠르게 연속 입력한다.
3. TOPAS 응답이 입력 속도를 따라오지 못하면 화면에는 긴 대기 구간과 로딩 표시가 남는다.
4. 시간이 지난 뒤 여러 날짜의 `AMADEUS AVAILABILITY` 결과가 한꺼번에 쌓인다.
5. 사용자는 날짜별 LJ915 클래스 잔여석을 스크롤하면서 확인해야 한다.

## v4.0 권장 워크플로우

1. 사용자가 TOPAS Entry 화면에서 첫 날짜 조회를 직접 실행한다. 예: `AN20JULDADICN/AZE594`.
2. 프로그램에서 `토파스조회` 탭을 선택하고 `토파스 조회하기`를 실행한다.
3. 팝업에서 AC1 실행 횟수를 입력한다.
4. 프로그램은 현재 첫 조회 원문을 기준 블록으로 보관하고, 이후 `AC1`을 10건씩 묶어 전송한다.
5. 각 묶음의 10개 응답 블록이 모두 출력된 것을 확인한 뒤 다음 묶음을 전송한다.
6. 묶음 응답은 최대 120초까지 기다리고, 일부만 확인되면 `N개 중 M개` 상태와 마지막 응답을 오류로 보고한다.
7. 조회 완료 후 흰 배경/검은 글자의 결과 팝업에 TOPAS 원문 전체를 표시하고, `전체 복사하기` 버튼으로 복사할 수 있게 한다.

## 설계 결정

- 기본 조회는 `첫날짜 사용자 직접조회 + 이후 AC1`을 우선한다.
- 사용자가 입력한 횟수와 상관없이 AC1 묶음 크기는 live test 결과 기준으로 10건 고정한다.
- 묶음 전체 응답 완료 감지 없이 다음 묶음으로 넘어가지 않는다.
- 화면에 남아 있는 오래된 스크롤백은 첫 조회 기준 블록 이후의 신규 응답만 인정해서 제외한다.
- 브라우저 제어는 Selenium 디버그 크롬 연결을 재사용하되, TOPAS 전용 설정을 `config.json`에 분리해서 추가한다.

## 1차 구현 범위

- `topas.availability`
  - 날짜 범위 -> TOPAS AN 명령 목록 생성
  - SellConnect 텍스트 응답 -> 구조화된 availability 데이터 파싱
  - Excel/CSV에 맞춘 flat row 변환
- `topas.pacing`
  - 다음 명령 전송 가능 여부 판단
  - 완료된 응답 블록 분리
- `gui.py`
  - `요금수정` / `토파스조회` 탭 분리
  - TOPAS 로그인 URL로 디버깅 크롬 실행
  - AC1 횟수 입력 팝업, 10건 묶음 전송, 전체 원문 결과 팝업
- `topas_cli.py`
  - GUI 통합 전 명령 생성/텍스트 파싱을 로컬에서 검증
- `tests/test_topas_availability.py`
  - 실제 영상에서 보인 LJ915 ICN-GUM 응답 형태 기반 단위 테스트

## 로컬 검증 명령

```powershell
python topas_cli.py commands 2026-07-01 2026-07-03 ICN GUM LJ 0915
python topas_cli.py commands 2026-07-01 2026-07-03 ICN GUM LJ 0915 --mode direct
python topas_cli.py parse sample_topas.txt --year 2026 --csv topas_availability.csv
python -m unittest tests.test_topas_availability
```

## live test 기록

- 시작 엔트리: `AN20JULDADICN/AZE594`
- 30회 조회 / 10건 묶음: 15.286초, 평균 0.510초/AC1, `AN21JUL~AN19AUG` 30개 응답 확인
- 30회 조회 / 30건 묶음: 15.445초, 평균 0.515초/AC1
- 결론: 30건 묶음 대비 속도 차이가 거의 없고, 안정성 측면에서 10건 묶음을 정식값으로 고정
