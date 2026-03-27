@echo off
setlocal

echo ============================================
echo  Team Dashboard Sync
echo ============================================

REM === CONFIG: update paths if needed =========
set "PY=C:\Users\jwagemd\OneDrive - Johnson Controls\Documents\MyPythonScripts\PyCharm\Scripts\python.exe"
set "PY_SCRIPT=C:\Users\jwagemd\OneDrive - Johnson Controls\Documents\Team Dashboard Files\teamDashboard_v2.py"
set "REPO_DIR=C:\Users\jwagemd\OneDrive - Johnson Controls\Documents\Team Dashboard Files\repo"
set "DASHBOARD_URL=https://YOUR-GITHUB-USERNAME.github.io/YOUR-REPO-NAME/techDashboard_v2.html"
set "BROWSER=msedge"
REM =============================================

echo.
echo [1/3] Running Python sync...
"%PY%" "%PY_SCRIPT%"
if errorlevel 1 (
    echo.
    echo ERROR: Python script failed. Dashboard NOT updated.
    pause
    exit /b 1
)

echo.
echo [2/3] Pushing data.json to GitHub...
cd /d "%REPO_DIR%"
git add data.json
git diff --cached --quiet
if errorlevel 1 (
    for /f "tokens=*" %%i in ('powershell -command "Get-Date -Format \"yyyy-MM-dd HH:mm\""') do set TIMESTAMP=%%i
    git commit -m "Dashboard sync %TIMESTAMP%"
    git push
    echo   Pushed successfully.
) else (
    echo   No changes to push ^(data unchanged since last sync^).
)

echo.
echo [3/3] Opening dashboard...
start "" "%BROWSER%" "%DASHBOARD_URL%"

echo.
echo Done.
pause
endlocal
