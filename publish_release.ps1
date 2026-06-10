<#
publish_release.ps1
  요금 업데이터 한 방 배포 스크립트.

  흐름:
    1) (기본) build_setup.ps1 실행해 현재 버전 설치본 빌드  (-SkipBuild 로 생략 가능)
    2) GitHub 릴리스 생성/업로드(--clobber)
    3) "업로드된 자산"을 다시 내려받아 SHA256 실측
    4) 루트 latest.json 의 version / download_url / sha256 을 그 값으로 동기화
    5) latest.json 커밋 & 푸시
    6) GitHub API(권위 소스)로 main 반영 확인

  목적: build_setup 이 해시를 release\latest.json 에만 쓰고 루트 latest.json 은 따로라
        새 버전마다 손으로 맞춰야 했고, 그 사이 해시 불일치로 자동 업데이트가 깨지던 문제를
        "업로드 자산 기준 해시 -> 루트 latest.json -> 푸시" 로 묶어 영구 차단.

  사용:
    .\publish_release.ps1                 # 빌드부터 전체
    .\publish_release.ps1 -SkipBuild      # 이미 빌드돼 있으면 업로드/푸시만
    .\publish_release.ps1 -Version v1.1.18
#>
param(
    [string]$Version = "",
    [switch]$SkipBuild,
    [string]$Repo = "KAKNAKIAK/naeil-erp-updater"
)

$ErrorActionPreference = "Stop"
$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ReleaseRoot = Join-Path $ProjectRoot "release"
$RootLatest  = Join-Path $ProjectRoot "latest.json"
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

# 0) 버전 결정: 인자 없으면 build_setup.ps1 의 $Version 사용
if (-not $Version) {
    $bs = Get-Content -LiteralPath (Join-Path $ProjectRoot "build_setup.ps1") -Raw
    if ($bs -match '\$Version\s*=\s*"([^"]+)"') { $Version = $Matches[1] }
}
if (-not $Version) { throw "버전을 확인할 수 없습니다. -Version v1.1.x 로 지정하세요." }
Write-Host "[publish] 대상 버전: $Version"

# 1) 빌드
if (-not $SkipBuild) {
    Write-Host "[publish] build_setup.ps1 실행..."
    & (Join-Path $ProjectRoot "build_setup.ps1")
    if ($LASTEXITCODE -ne 0) { throw "빌드 실패" }
}

$Setup = Join-Path $ReleaseRoot ("NaeilERPUpdater_Setup_{0}.exe" -f $Version)
if (-not (Test-Path -LiteralPath $Setup)) {
    throw "설치본이 없습니다: $Setup  (먼저 build_setup.ps1 실행 또는 -SkipBuild 해제)"
}

# 2) 릴리스 생성/업로드
# 자산 파일명은 실제 빌드 산출물($Setup)에서 떼온다 (확실)
$AssetName = Split-Path -Leaf $Setup
$ErrorActionPreference = "Continue"
gh release view $Version --repo $Repo *> $null
$releaseExists = ($LASTEXITCODE -eq 0)
$ErrorActionPreference = "Stop"

if ($releaseExists) {
    Write-Host "[publish] 기존 릴리스에 자산 업로드(덮어쓰기)..."
    gh release upload $Version $Setup --repo $Repo --clobber
} else {
    Write-Host "[publish] 새 릴리스 생성 + 자산 업로드..."
    gh release create $Version $Setup --repo $Repo --title $Version --notes ("Release {0}" -f $Version)
}
if ($LASTEXITCODE -ne 0) { throw "GitHub 릴리스 업로드 실패" }

# 3) 매니페스트 SHA256 = 방금 업로드한 그 파일(로컬 빌드 설치본)의 해시.
#    gh 는 바이트를 그대로 업로드하므로 "로컬 설치본 해시 == 사용자가 받는 자산 해시".
$DownloadUrl = "https://github.com/{0}/releases/download/{1}/{2}" -f $Repo, $Version, $AssetName
if ($DownloadUrl -notmatch '/NaeilERPUpdater_Setup_.+\.exe$') {
    throw "download_url 생성 오류(파일명 누락): '$DownloadUrl'  (AssetName='$AssetName', Version='$Version')"
}
$AssetHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $Setup).Hash
Write-Host ("[publish] 설치본 SHA256(로컬=업로드 자산): {0}" -f $AssetHash)

# (선택·비치명적) 업로드된 자산을 실제로 받아 한 번 더 대조. 실패해도 배포는 로컬 해시로 진행.
$PyExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (Test-Path -LiteralPath $PyExe) {
    $VerifyPy = Join-Path ([System.IO.Path]::GetTempPath()) ("verify_{0}.py" -f ([System.Guid]::NewGuid().ToString("N")))
@'
import urllib.request, hashlib, sys, time
url = sys.argv[1]
for _ in range(5):
    try:
        d = urllib.request.urlopen(urllib.request.Request(url, headers={"User-Agent": "x"}), timeout=180).read()
        sys.stdout.write(hashlib.sha256(d).hexdigest())
        sys.exit(0)
    except Exception:
        time.sleep(4)
sys.exit(2)
'@ | Set-Content -LiteralPath $VerifyPy -Encoding UTF8
    $remoteHash = (& $PyExe $VerifyPy $DownloadUrl)
    Remove-Item -LiteralPath $VerifyPy -Force -ErrorAction SilentlyContinue
    if ($remoteHash -match '^[0-9a-fA-F]{64}$') {
        $remoteHash = $remoteHash.Trim()
        if ($remoteHash -ne $AssetHash) {
            Write-Warning ("업로드 자산 해시({0})가 로컬({1})과 다릅니다. 업로드 자산 기준으로 작성합니다." -f $remoteHash, $AssetHash)
            $AssetHash = $remoteHash
        } else {
            Write-Host "[publish] 업로드 자산 재다운로드 대조 일치 OK"
        }
    } else {
        Write-Host "[publish] (참고) 재다운로드 대조 생략 - 로컬 해시 사용"
    }
}

# 4) 루트 latest.json 동기화 (notes 는 보존)
$j = Get-Content -LiteralPath $RootLatest -Raw -Encoding UTF8 | ConvertFrom-Json
$j.version = $Version
$j.download_url = $DownloadUrl
$j.sha256 = $AssetHash
# 푸시 직전 최종 가드 — 깨진 매니페스트는 절대 커밋/푸시하지 않는다
if ($j.download_url -notmatch '/NaeilERPUpdater_Setup_.+\.exe$') { throw "latest.json download_url 비정상: '$($j.download_url)'" }
if ($j.sha256 -notmatch '^[0-9a-fA-F]{64}$') { throw "latest.json sha256 비정상: '$($j.sha256)'" }
if ($j.version -ne $Version) { throw "latest.json version 불일치: '$($j.version)' != '$Version'" }
$RootLatestJson = $j | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText($RootLatest, $RootLatestJson, $Utf8NoBom)
Write-Host "[publish] 루트 latest.json 동기화 완료"

# 5) 커밋 & 푸시
Push-Location $ProjectRoot
try {
    git add latest.json
    git commit -m ("release: {0} 매니페스트 동기화(업로드 자산 해시 기준)" -f $Version)
    if ($LASTEXITCODE -ne 0) { Write-Host "[publish] 커밋할 변경이 없거나 커밋 실패(무시 가능)" }
    git push origin HEAD
    if ($LASTEXITCODE -ne 0) { throw "git push 실패" }
} finally {
    Pop-Location
}

# 6) 권위 소스 검증
Write-Host "[publish] GitHub API 로 main 반영 확인..."
$apiContent = gh api ("repos/{0}/contents/latest.json" -f $Repo) --jq '.content'
if ($LASTEXITCODE -eq 0 -and $apiContent) {
    $decoded = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String(($apiContent -replace '\s', '')))
    $decoded = $decoded.TrimStart([char]0xFEFF)   # UTF-8 BOM 제거(ConvertFrom-Json 오류 방지)
    $apiJson = $decoded | ConvertFrom-Json
    Write-Host ("[publish] main latest.json => version={0}  sha256={1}" -f $apiJson.version, $apiJson.sha256)
    if ($apiJson.sha256 -eq $AssetHash -and $apiJson.version -eq $Version) {
        Write-Host "[publish] OK: 릴리스 자산과 매니페스트 해시 일치 (배포 완료)"
    } else {
        Write-Warning "[publish] 주의: API 상의 값이 기대와 다릅니다. 확인 필요."
    }
} else {
    Write-Warning "[publish] API 확인 생략(네트워크/권한). 푸시는 완료됨."
}

Write-Host ""
Write-Host ("[publish] 완료: {0}" -f $Version)
Write-Host ("  설치본 : {0}" -f $Setup)
Write-Host ("  자산URL: {0}" -f $DownloadUrl)
Write-Host ("  SHA256 : {0}" -f $AssetHash)
Write-Host "  raw CDN 전파에 수 분 걸릴 수 있으나 main 은 즉시 정상입니다."
