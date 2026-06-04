@echo off
cd /d "%~dp0"
if "%1"=="clean" goto CLEAN

set HEADED=
if "%1"=="headed" set HEADED=--headed
if "%2"=="headed" set HEADED=--headed

call python --version >nul 2>nul
if errorlevel 1 (
    echo [FAIL] Python not found.
    pause
    exit /b 1
)

call python -c "import pptx" >nul 2>nul
if errorlevel 1 call pip install python-pptx

call python -c "import win32com.client" >nul 2>nul
if errorlevel 1 (
    echo [FAIL] pywin32 not found. Run: pip install pywin32
    pause
    exit /b 1
)

call python demo_test.py %HEADED%
pause
exit /b 0

:CLEAN
echo Cleaning test output...
del /q test_out*.pptx 2>nul
del /q test_out*.pdf 2>nul
del /q test_report.pptx 2>nul
del /q test_merged*.pptx 2>nul
del /q slide_*.png 2>nul
del /q test_img.png 2>nul
echo [OK] Cleanup done
pause
