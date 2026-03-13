@echo off
REM ===============================================
REM OpenClaw Environment Setup Script for New PC
REM ===============================================

echo Setting up OpenClaw Environment...
echo.

REM Create .openclaw directory in user profile if it doesn't exist
if not exist "%USERPROFILE%\.openclaw\" (
    mkdir "%USERPROFILE%\.openclaw"
)

REM Copy backed up config files to user profile
if exist "config\openclaw.json" (
    copy /Y "config\openclaw.json" "%USERPROFILE%\.openclaw\openclaw.json"
    echo Config restored.
) else (
    echo WARNING: config\openclaw.json not found in the backup.
)

REM Setup Windows Startup Shortcut
echo.
echo Setting up automatic startup...
powershell -Command "$WshShell = New-Object -ComObject WScript.Shell; $Shortcut = $WshShell.CreateShortcut('%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\OpenClaw_Dashboard.lnk'); $Shortcut.TargetPath = '%~dp0OpenClaw_Dashboard.bat'; $Shortcut.WorkingDirectory = '%~dp0'; $Shortcut.Save()"

echo.
echo ===============================================
echo Setup Complete! 
echo You can now use OpenClaw exactly like your previous PC.
echo ===============================================
pause
