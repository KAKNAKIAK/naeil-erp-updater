$ErrorActionPreference = "Stop"

$Version = "v5.0.13"

$ProjectRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$ReleaseRoot = Join-Path $ProjectRoot "release"
$ReleaseDir = Join-Path $ReleaseRoot "NaeilERPUpdater"
$SetupOut = Join-Path $ReleaseRoot "NaeilERPUpdater_Setup_$Version.exe"
$SetupAlias = Join-Path $ReleaseRoot "NaeilERPUpdater_Setup.exe"
$TempInstallerDir = Join-Path ([System.IO.Path]::GetTempPath()) "NaeilERPUpdaterInstaller2"
$Utf8NoBom = [System.Text.UTF8Encoding]::new($false)

function Reset-TempDir {
    param([Parameter(Mandatory = $true)][string]$Path)
    $TempRootFull = [System.IO.Path]::GetFullPath([System.IO.Path]::GetTempPath()).TrimEnd("\") + "\"
    $FullPath = [System.IO.Path]::GetFullPath($Path)
    if (-not $FullPath.StartsWith($TempRootFull, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing to modify a non-temp path: $FullPath"
    }
    if (Test-Path -LiteralPath $Path) {
        Remove-Item -LiteralPath $Path -Recurse -Force
    }
    New-Item -ItemType Directory -Force -Path $Path | Out-Null
}

$PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
if (-not (Test-Path -LiteralPath $PythonExe)) {
    python -m venv (Join-Path $ProjectRoot ".venv")
    $PythonExe = Join-Path $ProjectRoot ".venv\Scripts\python.exe"
}

Push-Location $ProjectRoot
try {
    Write-Host "[1/4] Checking PyInstaller..."
    $PreviousErrorActionPreference = $ErrorActionPreference
    $ErrorActionPreference = "Continue"
    & $PythonExe -m PyInstaller --version > $null 2> $null
    $PyInstallerCheckExitCode = $LASTEXITCODE
    $ErrorActionPreference = $PreviousErrorActionPreference

    if ($PyInstallerCheckExitCode -ne 0) {
        Write-Host "[1/4] Installing PyInstaller..."
        & $PythonExe -m pip install pyinstaller
        if ($LASTEXITCODE -ne 0) {
            throw "PyInstaller installation failed."
        }
    }

    Write-Host "[1/4] Building portable release (onedir)..."
    if (Test-Path -LiteralPath $ReleaseDir) {
        Remove-Item -LiteralPath $ReleaseDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    
    $TempPortableRoot = Join-Path ([System.IO.Path]::GetTempPath()) "NaeilERPUpdaterPortableBuild"
    $TempPortableDist = Join-Path $TempPortableRoot "dist"
    $TempPortableBuild = Join-Path $TempPortableRoot "build"
    Reset-TempDir -Path $TempPortableRoot

    & $PythonExe -m PyInstaller `
        --noconfirm `
        --clean `
        --distpath $TempPortableDist `
        --workpath $TempPortableBuild `
        (Join-Path $ProjectRoot "NaeilERPUpdater.spec")

    if ($LASTEXITCODE -ne 0) {
        throw "Portable app build failed."
    }

    $PortableSourceDir = Join-Path $TempPortableDist "NaeilERPUpdater"
    if (-not (Test-Path -LiteralPath (Join-Path $PortableSourceDir "NaeilERPUpdater.exe"))) {
        throw "Built portable release was not found: $PortableSourceDir"
    }

    Copy-Item -LiteralPath (Join-Path $ProjectRoot "config.json") -Destination $PortableSourceDir -Force
    Copy-Item -LiteralPath (Join-Path $ProjectRoot "chrome_debug.bat") -Destination $PortableSourceDir -Force

    # Copy to release folder for synchronization
    if (-not (Test-Path -LiteralPath $ReleaseRoot)) {
        New-Item -ItemType Directory -Force -Path $ReleaseRoot | Out-Null
    }
    Copy-Item -Path $PortableSourceDir -Destination $ReleaseRoot -Recurse -Force

    Write-Host "[2/4] Creating setup payload..."
    Reset-TempDir -Path $TempInstallerDir
    
    # Exclude temporary cache and profile directories to avoid file locks
    $PayloadSrc = Join-Path $TempInstallerDir "payload"
    New-Item -ItemType Directory -Force -Path $PayloadSrc | Out-Null
    Copy-Item -Path (Join-Path $PortableSourceDir "*") -Destination $PayloadSrc -Recurse -Force -Exclude "ChromeProfile", "logs"
    
    $PayloadZip = Join-Path $TempInstallerDir "payload.zip"
    Compress-Archive -Path (Join-Path $PayloadSrc "*") -DestinationPath $PayloadZip -Force

    Write-Host "[3/4] Building setup executable..."
    if (Test-Path -LiteralPath $SetupOut) {
        Remove-Item -LiteralPath $SetupOut -Force
    }
    & $PythonExe -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --windowed `
        --name "NaeilERPUpdater_Setup_$Version" `
        --icon (Join-Path $ProjectRoot "app_icon.ico") `
        --splash (Join-Path $ProjectRoot "install_splash.png") `
        --distpath $ReleaseRoot `
        --workpath (Join-Path $TempInstallerDir "build") `
        --specpath $TempInstallerDir `
        --add-data "$PayloadZip;." `
        (Join-Path $ProjectRoot "packaging\setup_installer.py")

    if ($LASTEXITCODE -ne 0 -or -not (Test-Path -LiteralPath $SetupOut)) {
        throw "Setup executable build failed."
    }

    Copy-Item -LiteralPath $SetupOut -Destination $SetupAlias -Force

    $LatestTemplate = Join-Path $ProjectRoot "latest.json"
    if (Test-Path -LiteralPath $LatestTemplate) {
        $LatestOut = Join-Path $ReleaseRoot "latest.json"
        $SetupHash = (Get-FileHash -Algorithm SHA256 -LiteralPath $SetupOut).Hash
        $DownloadUrl = "https://github.com/KAKNAKIAK/naeil-erp-updater/releases/download/$Version/NaeilERPUpdater_Setup_$Version.exe"
        $Latest = Get-Content -LiteralPath $LatestTemplate -Raw -Encoding UTF8 | ConvertFrom-Json
        $Latest.version = $Version
        $Latest.download_url = $DownloadUrl
        $Latest.sha256 = $SetupHash
        # guard: never write a broken manifest (missing filename / bad hash)
        if ($Latest.download_url -notmatch '\.exe$') { throw "latest.json download_url invalid (missing filename)" }
        if ($Latest.sha256 -notmatch '^[0-9a-fA-F]{64}$') { throw "latest.json sha256 invalid" }
        $Json = $Latest | ConvertTo-Json -Depth 10
        # 루트(푸시 대상)와 release 사본을 모두 동기화 -> 해시 불일치 영구 방지
        [System.IO.File]::WriteAllText($LatestOut, $Json, $Utf8NoBom)
        [System.IO.File]::WriteAllText($LatestTemplate, $Json, $Utf8NoBom)
        Write-Host "Latest manifest (release): $LatestOut"
        Write-Host "Latest manifest (root):    $LatestTemplate"
        Write-Host "Download URL: $DownloadUrl"
        Write-Host "Setup SHA256: $SetupHash"
    }

    Write-Host "[4/4] Done."
    Write-Host "Setup file: $SetupOut"
    Write-Host "Setup alias: $SetupAlias"
}
finally {
    Pop-Location
}
