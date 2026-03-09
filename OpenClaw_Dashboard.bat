@echo off
REM ===============================================
REM OpenClaw Dashboard & Gateway Launcher
REM ===============================================

echo ===============================================
netstat -ano | find "LISTENING" | find ":18792" >nul
if "%ERRORLEVEL%" equ "0" goto gateway_running

echo Starting OpenClaw Gateway Server (Gemini 3.0 Flash) - Port 18792 ...
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
start chrome "http://localhost:18792/?token=a06f59f3e119d52a2b6741bd290f440dc407bcb0931c6345"

echo.
echo Gateway is now running in its own window.
echo You can safely close THIS window, but leave the "OpenClaw Gateway Logs" window open.
echo.
pause
