@echo off
cd /d "%~dp0"
echo DEBUG: started
if "%1"=="clean" goto CLEAN
echo DEBUG: past clean check

echo ============================================
echo   PPT Editor - Windows Test Demo
echo ============================================
echo DEBUG: past header

echo DEBUG: about to check python
python --version
echo DEBUG: python check returned errorlevel %errorlevel%

echo DEBUG: about to check pptx
python -c "import pptx"
echo DEBUG: pptx check returned errorlevel %errorlevel%

echo DEBUG: about to check win32com
python -c "import win32com.client"
echo DEBUG: win32com check returned errorlevel %errorlevel%

echo === Step 1: Generate test PPTX ===
python gen_test.py
echo DEBUG: gen_test returned errorlevel %errorlevel%
echo [OK] test_report.pptx generated

echo === Step 2: COM Inspect ===
python pptx_editor_com.py test_report.pptx --inspect

echo.
echo All done!
pause
