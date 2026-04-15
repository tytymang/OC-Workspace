@echo off
set "CHROME=C:\Program Files\Google\Chrome\Application\chrome.exe"
set "CHROME_DEBUG_DIR=%USERPROFILE%\.openclaw\chrome-debug-profile"
set "OC_CMD=%APPDATA%\npm\openclaw.cmd"

REM -- Open loading page (same tab will redirect to Dashboard) --
start "" "%CHROME%" --remote-debugging-port=9222 --user-data-dir="%CHROME_DEBUG_DIR%" "%~dp0loading.html"

REM -- Wait for Gateway port then start Browser Relay --
set "WAITED=0"
:wait_loop
if %WAITED% geq 120 exit /b 1
for /f "tokens=5" %%a in ('netstat -aon 2^>^&1 ^| findstr ":18792" ^| findstr "LISTENING"') do goto gw_ready
ping -n 2 127.0.0.1 >"%TEMP%\_oc.log"
set /a WAITED+=1
goto wait_loop

:gw_ready
start "openclaw-browser" /min cmd /c ""%OC_CMD%" browser start"
exit
