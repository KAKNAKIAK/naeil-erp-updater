@echo off
title Chrome Debug Mode Launcher
chcp 65001 > nul

echo ============================================================
echo   Launching Chrome Browser in Debugging Mode
echo.
echo   [Instructions]
echo   1. Use this launcher from the updater app.
echo   2. ERP and TOPAS can use separate debug ports/profiles.
echo   3. Log in to the opened ERP or TOPAS page.
echo ============================================================
echo.

set "TARGET_URL=%~1"
if "%TARGET_URL%"=="" (
    set "TARGET_URL=https://erp.naeiltour.co.kr/erp/login"
)
set "DEBUG_PORT=%~2"
if "%DEBUG_PORT%"=="" (
    set "DEBUG_PORT=9222"
)
set "PROFILE_DIR=%~3"
if "%PROFILE_DIR%"=="" (
    set "PROFILE_DIR=ChromeProfile"
)
set "USER_DATA_DIR=%~dp0%PROFILE_DIR%"

:: 1. Verify Chrome Path
set "CHROME_PATH=%ProgramFiles%\Google\Chrome\Application\chrome.exe"
if not exist "%CHROME_PATH%" (
    set "CHROME_PATH=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe"
)
if not exist "%CHROME_PATH%" (
    set "CHROME_PATH=%LocalAppData%\Google\Chrome\Application\chrome.exe"
)

if not exist "%CHROME_PATH%" (
    echo [Error] Chrome installation path not found.
    echo Please install Google Chrome first.
    pause
    exit /b
)

:: 2. Launch Chrome
echo [System] Launching debugging Chrome browser...
echo [System] Target URL: %TARGET_URL%
echo [System] Debug Port: %DEBUG_PORT%
echo [System] Profile Dir: %USER_DATA_DIR%
start "" "%CHROME_PATH%" --remote-debugging-port=%DEBUG_PORT% --user-data-dir="%USER_DATA_DIR%" "%TARGET_URL%"
echo [Success] Chrome launched. You may close this command window.
timeout /t 3 > nul
exit
