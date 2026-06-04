@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM Clean mode
if "%1"=="clean" goto :clean

echo ============================================
echo   PPT Editor - Windows Local Test Demo
echo ============================================
echo.
echo [DEBUG] After header

REM Check Python
echo [DEBUG] Checking python...
python --version
echo [DEBUG] Python check done, errorlevel=%errorlevel%
if errorlevel 1 (
    echo [FAIL] Python not installed
    pause
    exit /b 1
)
echo [DEBUG] Python OK

REM Check python-pptx
echo [DEBUG] Checking python-pptx...
python -c "import pptx" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing python-pptx...
    pip install python-pptx
)
echo [DEBUG] python-pptx OK

echo.
echo ========== 1. Generate test PPTX ==========
python gen_test.py
if errorlevel 1 (
    echo [FAIL] Generation failed
    pause
    exit /b 1
)
echo [OK] test_report.pptx generated
echo.

echo ========== 2. Inspect structure ==========
python pptx_editor.py test_report.pptx --inspect
echo.

echo [DEBUG] Reached end of basic tests
pause
