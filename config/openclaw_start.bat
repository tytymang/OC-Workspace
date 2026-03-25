@echo off
setlocal enabledelayedexpansion
REM ===============================================
REM OpenClaw Start Script (Clean Identity & UI)
REM ===============================================

echo Initializing OpenClaw Environment...
if not exist "%USERPROFILE%\.openclaw" mkdir "%USERPROFILE%\.openclaw"

REM Sync Configuration & API Key
if exist "config\openclaw.json" copy /Y "config\openclaw.json" "%USERPROFILE%\.openclaw\openclaw.json" >nul
if exist "config\.env" (
    copy /Y "config\.env" "%USERPROFILE%\.openclaw\.env" >nul
    for /f "usebackq tokens=1,2 delims==" %%A in ("config\.env") do set "apikey=%%B"
)

REM Ensure Agent auth profiles
set "AGENT_DIR=%USERPROFILE%\.openclaw\agents\main\agent"
if not exist "%AGENT_DIR%" mkdir "%AGENT_DIR%"
(
echo {
echo   "profiles": {
echo     "google:default": {
echo       "provider": "google",
echo       "mode": "token",
echo       "token": "%apikey%"
echo     }
echo   }
echo }
) > "%AGENT_DIR!\auth-profiles.json"

echo.
echo Checking for existing Gateway on port 18792...
for /f "tokens=5" %%a in ('netstat -aon ^| findstr :18792 ^| findstr LISTENING') do (
    echo [INFO] Found existing Gateway. Restarting for clean state...
    taskkill /F /PID %%a >nul 2>&1
)

echo.
echo [STEP] Starting OpenClaw Gateway Server...
start "OpenClaw Gateway Logs" run_gateway.bat

echo Waiting for initialization (5s)...
timeout /t 5 /nobreak >nul

echo.
echo [STEP] Starting OpenClaw Browser...
start /min "" cmd /c "openclaw browser start >nul 2>&1"

echo.
echo [STEP] Launching Dashboard...
REM Use the base URL to allow the dashboard to initialize its own session state
start "" chrome "http://localhost:18792/?token=a06f59f3e119d52a2b6741bd290f440dc407bcb0931c6345"

echo.
echo ===============================================
echo Startup process completed.
echo [!] TIP: Please disable 'Google Translate' in your browser for this page.
echo ===============================================
echo This window will close in 5 seconds...
timeout /t 5 /nobreak >nul
exit
