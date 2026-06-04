@echo off
cd /d "%~dp0"
if "%1"=="clean" goto CLEAN

echo ============================================
echo   PPT Editor - Windows Test Demo
echo ============================================
echo.

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

echo.
echo === Step 1: Generate test PPTX ===
call python gen_test.py
if errorlevel 1 (
    echo [FAIL] gen_test.py failed
    pause
    exit /b 1
)
echo [OK] test_report.pptx generated
echo.
echo === Step 2: COM Inspect ===
call python pptx_editor_com.py test_report.pptx --inspect
echo.

echo === Step 3: Modify title text ===
call python pptx_editor_com.py test_report.pptx "put title as Demo Test OK" --output test_out_title.pptx
echo.

echo === Step 4: Font size ===
call python pptx_editor_com.py test_report.pptx "slide 1 title font size 40" --output test_out_fontsize.pptx
echo.

echo === Step 5: Bold ===
call python pptx_editor_com.py test_report.pptx "slide 1 title bold" --output test_out_bold.pptx
echo.

echo === Step 6: Color red ===
call python pptx_editor_com.py test_report.pptx "slide 1 title color red" --output test_out_color.pptx
echo.

echo === Step 7: Underline ===
call python pptx_editor_com.py test_report.pptx "slide 1 title underline" --output test_out_underline.pptx
echo.

echo === Step 8: Strikethrough ===
call python pptx_editor_com.py test_report.pptx "slide 1 title strikethrough" --output test_out_strikethrough.pptx
echo.

echo === Step 9: Center align ===
call python pptx_editor_com.py test_report.pptx "slide 1 title center align" --output test_out_align.pptx
echo.

echo === Step 10: Fill blue ===
call python pptx_editor_com.py test_report.pptx "slide 1 title background blue" --output test_out_fill.pptx
echo.

echo === Step 11: Border red ===
call python pptx_editor_com.py test_report.pptx "slide 1 title border red" --output test_out_border.pptx
echo.

echo === Step 12: Move to top-left ===
call python pptx_editor_com.py test_report.pptx "slide 1 title move to top left" --output test_out_move.pptx
echo.

echo === Step 13: Enlarge ===
call python pptx_editor_com.py test_report.pptx "slide 1 title enlarge" --output test_out_resize.pptx
echo.

echo === Step 14: Add textbox ===
call python pptx_editor_com.py test_report.pptx "add textbox with text Hello World" --output test_out_addtextbox.pptx
echo.

echo === Step 15: Animation fade in ===
call python pptx_editor_com.py test_report.pptx "slide 1 title fade in animation" --output test_out_animation.pptx
echo.

echo === Step 16: Transition fade ===
call python pptx_editor_com.py test_report.pptx "slide 1 transition fade" --output test_out_transition.pptx
echo.

echo === Step 17: Export PDF ===
call python pptx_editor_com.py test_report.pptx "export PDF" --output test_out_export.pptx
echo.

echo === Step 18: Export images ===
call python pptx_editor_com.py test_report.pptx --export-images
echo.

echo === Step 19: Add slide ===
call python pptx_editor_com.py test_report.pptx "add a slide" --output test_out_addslide.pptx
echo.

echo === Step 20: Delete slide 4 ===
call python pptx_editor_com.py test_report.pptx "delete slide 4" --output test_out_delslide.pptx
echo.

echo === Step 21: Move slide 3 to 1 ===
call python pptx_editor_com.py test_report.pptx "move slide 3 to position 1" --output test_out_moveslide.pptx
echo.

echo === Step 22: Table cell ===
call python pptx_editor_com.py test_report.pptx "table row 1 col 1 change to test data" --output test_out_cell.pptx
echo.

echo === Step 23: Add table row ===
call python pptx_editor_com.py test_report.pptx "table add a row" --output test_out_addrow.pptx
echo.

echo === Step 24: Add table column ===
call python pptx_editor_com.py test_report.pptx "table add a column" --output test_out_addcol.pptx
echo.

echo === Step 25: Remove animation ===
call python pptx_editor_com.py test_report.pptx "slide 1 remove animation" --output test_out_delanim.pptx
echo.

echo === Step 26: Delete table ===
call python pptx_editor_com.py test_report.pptx "delete table on slide 2" --output test_out_delete.pptx
echo.

echo.
echo ============================================
echo   All 26 tests completed!
echo ============================================
echo.
echo Generated files:
dir /b test_out_*.pptx 2>nul
echo.
pause
exit /b 0

:CLEAN
echo Cleaning test output...
del /q test_out_*.pptx 2>nul
del /q test_report.pptx 2>nul
del /q test_report.pdf 2>nul
del /q slide_*.png 2>nul
del /q pptx_editor_com.exe 2>nul
del /q pptx_editor_com.pdb 2>nul
echo [OK] Cleanup done
pause
