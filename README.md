# Naeil ERP Updater

내일투어 ERP 항공요금 자동 업데이트 도구입니다.

## Auto Update

앱은 시작 시 `config.json`의 `update_latest_url`을 확인합니다.

현재 연결된 manifest:

```text
https://raw.githubusercontent.com/KAKNAKIAK/naeil-erp-updater/main/latest.json
```

배포 절차:

1. `build_setup.ps1`로 설치본을 생성합니다.
2. `release/NaeilERPUpdater_Setup_<version>.exe`를 GitHub Release asset으로 업로드합니다.
3. 루트 `latest.json`의 `version`, `download_url`, `sha256`을 최신 설치본 기준으로 맞춥니다.
4. 변경된 `latest.json`을 `main` 브랜치에 push합니다.

앱은 `latest.json`의 `version`이 현재 `APP_VERSION`보다 높을 때만 업데이트를 제안합니다.
