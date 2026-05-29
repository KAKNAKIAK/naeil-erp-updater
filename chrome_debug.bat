@echo off
title Chrome Debug Mode Launcher
chcp 65001 > nul

echo ============================================================
echo   Launching Chrome Browser in Debugging Mode (Port: 9222)
echo.
echo   [Instructions]
echo   1. Close ALL existing normal Chrome browser windows first.
echo   2. Log in to Naeil Tour ERP in the newly opened Chrome window.
echo   3. Run the ERP updater program to inherit this active session.
echo ============================================================
echo.

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
start "" "%CHROME_PATH%" --remote-debugging-port=9222 --user-data-dir="%~dp0ChromeProfile" https://erp.naeiltour.co.kr/erp/login
echo [Success] Chrome launched. You may close this command window.
timeout /t 3 > nul
exit
