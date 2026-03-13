@echo off
REM Helper script to run OpenClaw Gateway cleanly
echo Starting Gateway...
call "%APPDATA%\npm\openclaw.cmd" gateway --port 18792
pause
