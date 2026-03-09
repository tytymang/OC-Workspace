@echo off
REM ===============================================
REM OpenClaw Dashboard & Gateway Launcher
REM ===============================================

echo ===============================================
netstat -ano | find "LISTENING" | find ":18792" >nul
if "%ERRORLEVEL%" equ "0" goto gateway_running

echo Starting OpenClaw Gateway Server (Gemini 2.5 Pro) - Port 18792 ...
echo A new command prompt window will open to show the Gateway logs.
start "OpenClaw Gateway Logs" run_gateway.bat
echo.
echo Waiting 5 seconds for Gateway to initialize...
timeout /t 5 /nobreak >nul
goto gateway_done

:gateway_running
echo Gateway is already running! Skipping Gateway startup...

:gateway_done
echo ===============================================

echo ===============================================
echo Launching OpenClaw Dashboard in Chrome...
echo ===============================================
set CHROME_PATH=
if exist "C:\Program Files\Google\Chrome\Application\chrome.exe" set CHROME_PATH="C:\Program Files\Google\Chrome\Application\chrome.exe"
if exist "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe" set CHROME_PATH="C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" set CHROME_PATH="%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"

if defined CHROME_PATH (
    start "" %CHROME_PATH% "http://localhost:18792/?token=a06f59f3e119d52a2b6741bd290f440dc407bcb0931c6345"
) else (
    echo [WARNING] Chrome NOT FOUND. Falling back to default browser...
    start "" "http://localhost:18792/?token=a06f59f3e119d52a2b6741bd290f440dc407bcb0931c6345"
)

echo.
echo Gateway is now running in its own window.
echo You can safely close THIS window, but leave the "OpenClaw Gateway Logs" window open.
echo.
pause
