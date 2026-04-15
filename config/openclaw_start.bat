@echo off
setlocal
cd /d "%~dp0"

REM -- Launch post-start in background --
start /min "" cmd /c ""%~dp0openclaw_post_start.bat""

REM -- Run Gateway in this window --
title openclaw-gateway
echo Starting OpenClaw Gateway...
call "%APPDATA%\npm\openclaw.cmd" gateway --port 18792 --force
pause
