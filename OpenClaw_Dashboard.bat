@echo off
REM ===============================================
REM OpenClaw Dashboard & Gateway Launcher
REM ===============================================

echo ===============================================
echo Starting OpenClaw Gateway Server (Port 18792)...
echo A new command prompt window will open to show the Gateway logs.
echo ===============================================
start "OpenClaw Gateway Logs" cmd /k "openclaw gateway --port 18792"

echo.
echo Waiting 5 seconds for Gateway to initialize...
timeout /t 5 /nobreak >nul

echo ===============================================
echo Launching OpenClaw Dashboard...
echo ===============================================
start chrome "http://localhost:18792/"

echo.
echo Gateway is now running in its own window.
echo You can safely close THIS window, but leave the "OpenClaw Gateway Logs" window open.
echo.
pause
