@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"

REM Clean mode
if "%1"=="clean" goto :clean

echo ============================================
echo   PPT Editor - Windows Local Test Demo
echo ============================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [FAIL] Python not installed
    pause
    exit /b 1
)

REM Check python-pptx
python -c "import pptx" >nul 2>&1
if errorlevel 1 (
    echo [INFO] Installing python-pptx...
    pip install python-pptx
)

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

echo ========== 3. Test: modify title ==========
python pptx_editor.py test_report.pptx "put title as Demo Test OK" --output test_out_title.pptx
echo.

echo ========== 4. Test: font size ==========
python pptx_editor.py test_report.pptx "slide 1 title font size 40" --output test_out_fontsize.pptx
echo.

echo ========== 5. Test: bold ==========
python pptx_editor.py test_report.pptx "slide 1 title bold" --output test_out_bold.pptx
echo.

echo ========== 6. Test: color ==========
python pptx_editor.py test_report.pptx "slide 1 title color red" --output test_out_color.pptx
echo.

echo ========== 7. Test: delete element ==========
python pptx_editor.py test_report.pptx "delete table on slide 2" --output test_out_delete.pptx
echo.

REM COM tests (require Office)
echo ========== 8. COM Tests (require Office) ==========
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo [WARN] pywin32 not installed, skipping COM tests
    echo        Install: pip install pywin32
    goto :done
)

echo --- 8a. COM inspect ---
python pptx_editor_com.py test_report.pptx --inspect
echo.

echo --- 8b. COM animation ---
python pptx_editor_com.py test_report.pptx "slide 1 title fade in animation" --output test_out_animation.pptx
echo.

echo --- 8c. COM transition ---
python pptx_editor_com.py test_report.pptx "slide 1 transition fade" --output test_out_transition.pptx
echo.

echo --- 8d. COM export PDF ---
python pptx_editor_com.py test_report.pptx "export PDF" --output test_out_export.pptx
echo.

echo --- 8e. COM export images ---
python pptx_editor_com.py test_report.pptx --export-images
echo.

echo.
echo ========== 9. Extended Tests ==========
echo.

echo --- 9a. Add textbox ---
python pptx_editor_com.py test_report.pptx "add textbox with text Hello World" --output test_out_addtextbox.pptx
echo.

echo --- 9b. Underline ---
python pptx_editor_com.py test_report.pptx "slide 1 title underline" --output test_out_underline.pptx
echo.

echo --- 9c. Strikethrough ---
python pptx_editor_com.py test_report.pptx "slide 1 title strikethrough" --output test_out_strikethrough.pptx
echo.

echo --- 9d. Center align ---
python pptx_editor_com.py test_report.pptx "slide 1 title center align" --output test_out_align.pptx
echo.

echo --- 9e. Fill color ---
python pptx_editor_com.py test_report.pptx "slide 1 title background blue" --output test_out_fill.pptx
echo.

echo --- 9f. Border ---
python pptx_editor_com.py test_report.pptx "slide 1 title border red" --output test_out_border.pptx
echo.

echo --- 9g. Move element ---
python pptx_editor_com.py test_report.pptx "slide 1 title move to top left" --output test_out_move.pptx
echo.

echo --- 9h. Resize element ---
python pptx_editor_com.py test_report.pptx "slide 1 title enlarge" --output test_out_resize.pptx
echo.

echo --- 9i. Add slide ---
python pptx_editor_com.py test_report.pptx "add a slide" --output test_out_addslide.pptx
echo.

echo --- 9j. Delete slide ---
python pptx_editor_com.py test_report.pptx "delete slide 4" --output test_out_delslide.pptx
echo.

echo --- 9k. Move slide ---
python pptx_editor_com.py test_report.pptx "move slide 3 to position 1" --output test_out_moveslide.pptx
echo.

echo --- 9l. Modify table cell ---
python pptx_editor_com.py test_report.pptx "table row 1 col 1 change to test data" --output test_out_cell.pptx
echo.

echo --- 9m. Add table row ---
python pptx_editor_com.py test_report.pptx "table add a row" --output test_out_addrow.pptx
echo.

echo --- 9n. Add table column ---
python pptx_editor_com.py test_report.pptx "table add a column" --output test_out_addcol.pptx
echo.

echo --- 9o. Insert picture ---
python pptx_editor_com.py test_report.pptx "slide 1 insert picture test_img.png" --output test_out_picture.pptx
echo.

echo --- 9p. Remove animation ---
python pptx_editor_com.py test_report.pptx "slide 1 remove animation" --output test_out_delanim.pptx
echo.

echo.
echo ========== 10. C# .NET COM Tests ==========
echo.

REM Find csc.exe
set CSC=
if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if defined CSC (
    echo [INFO] Compiling C# COM version...

    REM Find Office Interop DLL
    set PIA_PP=
    set PIA_OFFICE=
    for /f "delims=" %%d in ('dir /s /b "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint" 2^>nul') do (
        for /f "delims=" %%f in ('dir /b /s "%%d\Microsoft.Office.Interop.PowerPoint.dll" 2^>nul') do set PIA_PP=%%f
    )
    for /f "delims=" %%d in ('dir /s /b "C:\Windows\assembly\GAC_MSIL\office" 2^>nul') do (
        for /f "delims=" %%f in ('dir /b /s "%%d\office.dll" 2^>nul') do set PIA_OFFICE=%%f
    )

    if defined PIA_PP if defined PIA_OFFICE (
        %CSC% /nologo /out:pptx_editor_com.exe /reference:"!PIA_PP!" /reference:"!PIA_OFFICE!" pptx_editor_com.cs >nul 2>&1
        if exist pptx_editor_com.exe (
            echo [OK] Compiled successfully
            echo.
            echo --- 10a. C# inspect ---
            pptx_editor_com.exe test_report.pptx --inspect
            echo.
            echo --- 10b. C# modify title ---
            pptx_editor_com.exe test_report.pptx "change title to CSharp Test OK" --output test_out_cs_title.pptx
            echo.
            echo --- 10c. C# bold ---
            pptx_editor_com.exe test_report.pptx "slide 1 title bold" --output test_out_cs_bold.pptx
            echo.
            echo --- 10d. C# animation ---
            pptx_editor_com.exe test_report.pptx "slide 1 title fade in animation" --output test_out_cs_anim.pptx
            echo.
            echo --- 10e. C# transition ---
            pptx_editor_com.exe test_report.pptx "slide 1 transition fade" --output test_out_cs_trans.pptx
            echo.
            echo --- 10f. C# add textbox ---
            pptx_editor_com.exe test_report.pptx "add textbox with text CSharp Test" --output test_out_cs_textbox.pptx
            echo.
            echo --- 10g. C# center align ---
            pptx_editor_com.exe test_report.pptx "slide 1 title center align" --output test_out_cs_align.pptx
            echo.
            echo --- 10h. C# fill color ---
            pptx_editor_com.exe test_report.pptx "slide 1 title background blue" --output test_out_cs_fill.pptx
            echo.
        ) else (
            echo [WARN] C# compilation failed
            %CSC% /nologo /out:pptx_editor_com.exe /reference:"!PIA_PP!" /reference:"!PIA_OFFICE!" pptx_editor_com.cs
        )
    ) else (
        echo [WARN] Office Interop DLL not found, skipping C# tests
        echo        Need Office PIA installed
    )
) else (
    echo [WARN] csc.exe not found, skipping C# tests
    echo        Need .NET Framework 4.x
)
echo.

:done
echo.
echo ============================================
echo   [OK] All tests completed!
echo ============================================
echo.
echo Generated files:
dir /b *.pptx *.pdf *.png 2>nul
echo.

echo.
echo Tip: Clean test output with: demo_test.bat clean
pause
exit /b 0

:clean
echo Cleaning test output...
set count=0
for %%f in (*_modified.pptx test_out_*.pptx) do (if exist "%%f" (del "%%f" & echo   Deleted %%f & set /a count+=1))
for %%f in (*_structure.json) do (if exist "%%f" (del "%%f" & echo   Deleted %%f & set /a count+=1))
for %%f in (test_report.pptx) do (if exist "%%f" (del "%%f" & echo   Deleted %%f & set /a count+=1))
for %%f in (test_report.pdf) do (if exist "%%f" (del "%%f" & echo   Deleted %%f & set /a count+=1))
for %%f in (slide_*.png test_img.png) do (if exist "%%f" (del "%%f" & echo   Deleted %%f & set /a count+=1))
for %%f in (pptx_editor_com.exe pptx_editor_com.pdb) do (if exist "%%f" (del "%%f" & echo   Deleted %%f & set /a count+=1))
echo.
echo [OK] Cleanup done
pause
