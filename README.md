# Naeil ERP Updater

내일투어 ERP 항공요금 자동 업데이트 및 TOPAS AC1 연속 조회 도구입니다.

## v4.0.0

- `요금수정` / `토파스조회` 화면을 탭으로 분리했습니다.
- `브라우저 켜기`는 선택된 탭에 따라 ERP 로그인 또는 TOPAS 로그인 화면을 엽니다.
- TOPAS 조회는 사용자가 첫 날짜를 직접 조회한 뒤, `AC1`을 10건씩 묶어 자동 실행합니다.
- 조회 완료 후 TOPAS 원문 전체를 복사 가능한 팝업으로 제공합니다.

## Auto Update

앱은 시작 시 `config.json`의 `update_latest_url`을 확인합니다.

현재 연결된 manifest:

```text
https://api.github.com/repos/KAKNAKIAK/naeil-erp-updater/contents/latest.json?ref=main
```

배포 절차:

1. `build_setup.ps1`로 설치본을 생성합니다.
2. `release/NaeilERPUpdater_Setup_<version>.exe`를 GitHub Release asset으로 업로드합니다.
3. 루트 `latest.json`의 `version`, `download_url`, `sha256`을 최신 설치본 기준으로 맞춥니다.
4. 변경된 `latest.json`을 `main` 브랜치에 push합니다.

앱은 `latest.json`의 `version`이 현재 `APP_VERSION`보다 높을 때만 업데이트를 제안합니다.
