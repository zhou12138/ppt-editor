@echo off
cd /d "%~dp0"
set "SCRIPTS=%~dp0skills\pptx-local-offline\scripts"
if /I "%1"=="clean" goto CLEAN

set "PYTHONUTF8=1"
chcp 65001 >nul

call python -X utf8 --version >nul 2>nul
if errorlevel 1 (
    echo [FAIL] Python not found.
    pause
    exit /b 1
)

if not exist "%SCRIPTS%\pptx_editor.py" (
    echo [FAIL] Missing file: %SCRIPTS%\pptx_editor.py
    pause
    exit /b 1
)

if not exist "%SCRIPTS%\pptx_editor_com.py" (
    echo [FAIL] Missing file: %SCRIPTS%\pptx_editor_com.py
    pause
    exit /b 1
)

if not exist "%SCRIPTS%\pptx_editor_llm.py" (
    echo [FAIL] Missing file: %SCRIPTS%\pptx_editor_llm.py
    pause
    exit /b 1
)

echo ===================================================
echo   PPT Editor Skill - Script Smoke Test
echo   Target: skills\pptx-local-offline\scripts
echo ===================================================
echo.

echo [1/4] Checking Python syntax...
call python -X utf8 -m py_compile "%SCRIPTS%\pptx_editor.py" "%SCRIPTS%\pptx_editor_com.py" "%SCRIPTS%\pptx_editor_llm.py"
if errorlevel 1 goto FAIL
echo [OK] Syntax check passed.
echo.

echo [2/4] Starting python-pptx script without arguments...
call python -X utf8 "%SCRIPTS%\pptx_editor.py" >nul
if errorlevel 1 goto FAIL
echo [OK] pptx_editor.py startup check passed.
echo.

echo [3/4] Starting COM script without arguments...
call python -X utf8 "%SCRIPTS%\pptx_editor_com.py" >nul
if errorlevel 1 goto FAIL
echo [OK] pptx_editor_com.py startup check passed.
echo.

echo [4/4] Checking LLM wrapper help output...
call python -X utf8 "%SCRIPTS%\pptx_editor_llm.py" --help >nul
if errorlevel 1 goto FAIL
echo [OK] pptx_editor_llm.py help check passed.
echo.

echo All skill script smoke tests passed.
pause
exit /b 0

:CLEAN
echo ===================================================
echo   Cleaning skill script test artifacts...
echo ===================================================
echo.
if exist "%SCRIPTS%\__pycache__" rmdir /s /q "%SCRIPTS%\__pycache__"
echo Done. Removed skill script __pycache__.
pause
exit /b 0

:FAIL
echo.
echo [FAIL] Skill script smoke test failed.
pause
exit /b 1