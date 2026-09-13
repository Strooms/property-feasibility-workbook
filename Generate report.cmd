@echo off
setlocal
cd /d "%~dp0"

set "PY=.venv\Scripts\python.exe"
if not exist "%PY%" (
    echo Could not find %PY%
    echo Create the virtual environment first - see README.md
    pause
    exit /b 1
)

echo.
echo   Property feasibility report
echo   ---------------------------
echo     1^) Residential
echo     2^) Townhouse
echo.
set "SHEET="
set /p CHOICE="  Which deal type? [1] "
if "%CHOICE%"=="2" (set "SHEET=Townhouse") else (set "SHEET=Residential")

echo.
"%PY%" build_report.py --sheet %SHEET% --pdf --open
if errorlevel 1 (
    echo.
    echo   Generation failed. The messages above say why.
    pause
    exit /b 1
)
endlocal
