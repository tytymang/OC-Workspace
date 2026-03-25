@echo off
chcp 65001 >nul
echo ==========================================
echo   OpenClaw Google API Key Change Tool
echo ==========================================
echo.

set /p NEW_KEY="New Google API Key: "

if "%NEW_KEY%"=="" (
    echo ERROR: No key entered.
    pause
    exit /b 1
)

echo.
echo Updating 3 files...

:: 1. auth-profiles.json (PRIMARY - OpenClaw actually reads this)
echo {"version":1,"profiles":{"google:default":{"provider":"google","mode":"token","token":"%NEW_KEY%","type":"token"}},"usageStats":{"google:default":{"errorCount":0,"lastUsed":0}}} > "C:\Users\307984\.openclaw\agents\main\agent\auth-profiles.json"
echo [1/3] auth-profiles.json - DONE

:: 2. .openclaw/.env
echo GOOGLE_API_KEY=%NEW_KEY%> "C:\Users\307984\.openclaw\.env"
echo [2/3] .openclaw/.env - DONE

:: 3. workspace/config/.env
echo GOOGLE_API_KEY=%NEW_KEY%> "C:\Users\307984\.openclaw\workspace\config\.env"
echo [3/3] workspace/config/.env - DONE

echo.
echo ==========================================
echo   ALL DONE! Now restart OpenClaw.
echo ==========================================
pause
