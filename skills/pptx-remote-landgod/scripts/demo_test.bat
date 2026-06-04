@echo off
cd /d "%~dp0"
if "%1"=="clean" goto CLEAN
set VIS=False
if "%1"=="headed" set VIS=True

call python --version >nul 2>nul
if errorlevel 1 (
    echo [FAIL] Python not found.
    pause
    exit /b 1
)

echo ===================================================
echo   PPT Editor COM - Demo Test
echo   Mode: %VIS%
echo ===================================================
echo.

if "%VIS%"=="True" (
    call python demo_test.py --headed
) else (
    call python demo_test.py
)

echo.
pause
exit /b 0

:CLEAN
echo ===================================================
echo   Cleaning demo test artifacts...
echo ===================================================
echo.
del /q test_report.pptx 2>nul
del /q test_report_modified.pptx 2>nul
del /q test_report_structure.json 2>nul
del /q test_img.png 2>nul
del /q test_bg.png 2>nul
del /q demo_output.pptx 2>nul
del /q demo_output.pdf 2>nul
del /q demo_final.pptx 2>nul
del /q demo_merged.pptx 2>nul
del /q demo_slide1.png 2>nul
del /q *_modified.pptx 2>nul
del /q *_structure.json 2>nul
echo.
echo Done. All test artifacts removed.
pause
exit /b 0
