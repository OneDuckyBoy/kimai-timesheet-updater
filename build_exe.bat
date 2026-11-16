@echo off
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   Building timesheet-updater.exe (PyInstaller)
echo ============================================
echo.

:: Prefer a local virtualenv if present, else the system python.
if exist ".venv\Scripts\python.exe" (
    set "PY=.venv\Scripts\python.exe"
) else (
    set "PY=python"
)

echo Installing build + runtime dependencies...
%PY% -m pip install --upgrade pyinstaller openpyxl xlrd
if errorlevel 1 (
    echo [ERROR] pip install failed.
    pause
    exit /b 1
)

echo.
echo Building...
%PY% -m PyInstaller --noconfirm --distpath dist --workpath build timesheet-updater.spec

if exist "dist\timesheet-updater.exe" (
    echo.
    echo [DONE] Built dist\timesheet-updater.exe
) else (
    echo.
    echo [ERROR] Build failed - no dist\timesheet-updater.exe produced.
)

pause
