@echo off
title Chrome Debug Mode Launcher
chcp 65001 > nul

echo ============================================================
echo   Launching Chrome Browser in Debugging Mode (Port: 9222)
echo.
echo   [Instructions]
echo   1. Close ALL existing normal Chrome browser windows first.
echo   2. Log in to the opened ERP or TOPAS page.
echo   3. Run the updater program to inherit this active session.
echo ============================================================
echo.

set "TARGET_URL=%~1"
if "%TARGET_URL%"=="" (
    set "TARGET_URL=https://erp.naeiltour.co.kr/erp/login"
)

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
start "" "%CHROME_PATH%" --remote-debugging-port=9222 --user-data-dir="%~dp0ChromeProfile" "%TARGET_URL%"
echo [Success] Chrome launched. You may close this command window.
timeout /t 3 > nul
exit
