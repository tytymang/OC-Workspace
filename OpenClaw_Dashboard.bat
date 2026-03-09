@echo off
REM ===============================================
REM OpenClaw Dashboard & Gateway Launcher
REM ===============================================

echo ===============================================
netstat -ano | find "LISTENING" | find ":18792" >nul
if %ERRORLEVEL% equ 0 (
    echo Gateway is already running! Skipping Gateway startup...
) else (
    echo Starting OpenClaw Gateway Server - Port 18792 ...
    echo A new command prompt window will open to show the Gateway logs.
    start "OpenClaw Gateway Logs" cmd /k "%APPDATA%\npm\openclaw.cmd gateway --port 18792"
    echo.
    echo Waiting 5 seconds for Gateway to initialize...
    timeout /t 5 /nobreak >nul
)
echo ===============================================

echo ===============================================
echo Launching OpenClaw Dashboard...
echo ===============================================
start chrome "http://localhost:18792/"

echo.
echo Gateway is now running in its own window.
echo You can safely close THIS window, but leave the "OpenClaw Gateway Logs" window open.
echo.
pause
