@echo off
REM ===============================================
REM OpenClaw Workspace Sync Script
REM ===============================================

echo Syncing OpenClaw configurations to Github...
echo.

REM Ensure config directory exists
if not exist "config" (
    mkdir config
)

REM Backup current user configurations
echo Backing up %USERPROFILE%\.openclaw config files...
copy /Y "%USERPROFILE%\.openclaw\openclaw.json" "config\openclaw.json"
if exist "%USERPROFILE%\.openclaw\.env" (
    copy /Y "%USERPROFILE%\.openclaw\.env" "config\.env"
)

echo.
echo Committing to GitHub...
git add .
set datetime=%date% %time%
git commit -m "Auto sync: %datetime%"

echo Pushing changes...
git push -u origin main

echo.
echo ===============================================
echo Sync Complete!
echo ===============================================
pause
