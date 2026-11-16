@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
cd /d "%~dp0"

echo ============================================
echo   Kimai Timesheet Updater
echo ============================================
echo.

:: Find Excel file in input\ (xlsx or xls)
set "EXCEL="
set "COUNT=0"
for /f "delims=" %%f in ('dir /b input\ 2^>nul') do (
    set "ext=%%~xf"
    if /i "!ext!"==".xlsx" (
        set /a COUNT+=1
        set "EXCEL=input\%%f"
    )
    if /i "!ext!"==".xls" (
        set /a COUNT+=1
        set "EXCEL=input\%%f"
    )
)
if !COUNT! gtr 1 (
    echo [ERROR] Multiple Excel files found in input\. Keep only one!
    pause
    exit /b 1
)
if !COUNT! equ 0 (
    echo [ERROR] No Excel file found in input\
    echo Put your Excel template ^(.xlsx or .xls^) in the input\ folder.
    pause
    exit /b 1
)

:: Find CSV file in input\
set "CSV="
set "COUNT=0"
for /f "delims=" %%f in ('dir /b input\ 2^>nul') do (
    set "ext=%%~xf"
    if /i "!ext!"==".csv" (
        set /a COUNT+=1
        set "CSV=input\%%f"
    )
)
if !COUNT! gtr 1 (
    echo [ERROR] Multiple CSV files found in input\. Keep only one!
    pause
    exit /b 1
)
if !COUNT! equ 0 (
    echo [ERROR] No CSV file found in input\
    echo Put your Kimai CSV export in the input\ folder.
    pause
    exit /b 1
)

echo [OK] Excel: %EXCEL%
echo [OK] CSV:   %CSV%
echo.

:: Run the updater (portable .exe - no Python required)
:: Output goes to output\ with the template's name + the new month.
internal\dist\timesheet-updater.exe "%EXCEL%" "%CSV%" --outdir output

if errorlevel 1 (
    echo.
    echo [ERROR] Script failed! See above for details.
) else (
    echo.
    echo [DONE] Output saved to output\ - see filename with the new month
)

echo.
pause