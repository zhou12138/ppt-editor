@echo off
chcp 65001 >nul 2>&1
cd /d "%~dp0"

REM 清理模式
if "%1"=="clean" goto :clean

echo ============================================
echo   PPT Editor - Windows 本地测试 Demo
echo ============================================
echo.

REM 检查依赖
python --version >nul 2>&1
if errorlevel 1 (
    echo ❌ Python 未安装
    pause
    exit /b 1
)

REM 检查 python-pptx
python -c "import pptx" >nul 2>&1
if errorlevel 1 (
    echo 📦 安装 python-pptx...
    pip install python-pptx
)

echo.
echo ========== 1. 生成测试 PPT ==========
python gen_test.py
if errorlevel 1 (
    echo ❌ 生成失败
    pause
    exit /b 1
)
echo ✅ test_report.pptx 已生成
echo.

echo ========== 2. 查看结构 ==========
python pptx_editor.py test_report.pptx --inspect
echo.

echo ========== 3. 测试：修改标题 ==========
python pptx_editor.py test_report.pptx "把标题改成「Demo 测试成功」" --output test_out_title.pptx
echo.

echo ========== 4. 测试：修改字号 ==========
python pptx_editor.py test_report.pptx "第1页标题字号改成40" --output test_out_fontsize.pptx
echo.

echo ========== 5. 测试：加粗 ==========
python pptx_editor.py test_report.pptx "第1页标题加粗" --output test_out_bold.pptx
echo.

echo ========== 6. 测试：改颜色 ==========
python pptx_editor.py test_report.pptx "第1页标题颜色改成红色" --output test_out_color.pptx
echo.

echo ========== 7. 测试：删除元素 ==========
python pptx_editor.py test_report.pptx "删除第2页的表格" --output test_out_delete.pptx
echo.

REM COM 版测试（需要 Office）
echo ========== 8. COM 版测试（需要 Office）==========
python -c "import win32com.client" >nul 2>&1
if errorlevel 1 (
    echo ⚠️  pywin32 未安装，跳过 COM 测试
    echo    安装: pip install pywin32
    goto :done
)

echo --- 8a. COM 查看结构 ---
python pptx_editor_com.py test_report.pptx --inspect
echo.

echo --- 8b. COM 添加动画 ---
python pptx_editor_com.py test_report.pptx "给第1页标题添加动画淡入" --output test_out_animation.pptx
echo.

echo --- 8c. COM 切换效果 ---
python pptx_editor_com.py test_report.pptx "第1页切换效果淡化" --output test_out_transition.pptx
echo.

echo --- 8d. COM 导出 PDF ---
python pptx_editor_com.py test_report.pptx "导出PDF" --output test_out_export.pptx
echo.

echo --- 8e. COM 导出图片 ---
python pptx_editor_com.py test_report.pptx --export-images
echo.

echo.
echo ========== 9. 新增场景测试 ==========
echo.

echo --- 9a. 添加文本框 ---
python pptx_editor_com.py test_report.pptx "添加文本框 内容是Hello World" --output test_out_addtextbox.pptx
echo.

echo --- 9b. 下划线 ---
python pptx_editor_com.py test_report.pptx "第1页标题下划线" --output test_out_underline.pptx
echo.

echo --- 9c. 删除线 ---
python pptx_editor_com.py test_report.pptx "第1页标题删除线" --output test_out_strikethrough.pptx
echo.

echo --- 9d. 居中对齐 ---
python pptx_editor_com.py test_report.pptx "第1页标题居中对齐" --output test_out_align.pptx
echo.

echo --- 9e. 背景填充 ---
python pptx_editor_com.py test_report.pptx "第1页标题背景改成蓝色" --output test_out_fill.pptx
echo.

echo --- 9f. 边框 ---
python pptx_editor_com.py test_report.pptx "第1页标题边框改成红色" --output test_out_border.pptx
echo.

echo --- 9g. 移动元素 ---
python pptx_editor_com.py test_report.pptx "第1页标题移动到左上" --output test_out_move.pptx
echo.

echo --- 9h. 放大元素 ---
python pptx_editor_com.py test_report.pptx "第1页标题放大" --output test_out_resize.pptx
echo.

echo --- 9i. 添加页面 ---
python pptx_editor_com.py test_report.pptx "添加一页" --output test_out_addslide.pptx
echo.

echo --- 9j. 删除页面 ---
python pptx_editor_com.py test_report.pptx "删除第4页" --output test_out_delslide.pptx
echo.

echo --- 9k. 移动页面 ---
python pptx_editor_com.py test_report.pptx "第3页移到第1页" --output test_out_moveslide.pptx
echo.

echo --- 9l. 修改表格单元格 ---
python pptx_editor_com.py test_report.pptx "表格第1行第1列改成测试数据" --output test_out_cell.pptx
echo.

echo --- 9m. 表格添加一行 ---
python pptx_editor_com.py test_report.pptx "表格添加一行" --output test_out_addrow.pptx
echo.

echo --- 9n. 表格添加一列 ---
python pptx_editor_com.py test_report.pptx "表格添加一列" --output test_out_addcol.pptx
echo.

echo --- 9o. 插入图片 ---
python pptx_editor_com.py test_report.pptx "第1页插入图片 test_img.png" --output test_out_picture.pptx
echo.

echo --- 9p. 删除动画 ---
python pptx_editor_com.py test_report.pptx "第1页删除动画" --output test_out_delanim.pptx
echo.

echo.
echo ========== 10. C# .NET COM 测试 ==========
echo.

REM 查找 csc.exe
set CSC=
if exist "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe" set CSC=C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe
if defined CSC (
    echo 🔧 编译 C# COM 版...
    
    REM 查找 Office Interop DLL
    set PIA_PP=
    set PIA_OFFICE=
    for /f "delims=" %%d in ('dir /s /b "C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.PowerPoint" 2^>nul') do (
        for /f "delims=" %%f in ('dir /b /s "%%d\Microsoft.Office.Interop.PowerPoint.dll" 2^>nul') do set PIA_PP=%%f
    )
    for /f "delims=" %%d in ('dir /s /b "C:\Windows\assembly\GAC_MSIL\office" 2^>nul') do (
        for /f "delims=" %%f in ('dir /b /s "%%d\office.dll" 2^>nul') do set PIA_OFFICE=%%f
    )
    
    if defined PIA_PP if defined PIA_OFFICE (
        %CSC% /nologo /out:pptx_editor_com.exe /reference:"%PIA_PP%" /reference:"%PIA_OFFICE%" pptx_editor_com.cs >nul 2>&1
        if exist pptx_editor_com.exe (
            echo ✅ 编译成功
            echo.
            echo --- 10a. C# 查看结构 ---
            pptx_editor_com.exe test_report.pptx --inspect
            echo.
            echo --- 10b. C# 修改标题 ---
            pptx_editor_com.exe test_report.pptx "把标题改成「C# 测试成功」" --output test_out_cs_title.pptx
            echo.
            echo --- 10c. C# 加粗+颜色 ---
            pptx_editor_com.exe test_report.pptx "第1页标题加粗" --output test_out_cs_bold.pptx
            echo.
            echo --- 10d. C# 动画 ---
            pptx_editor_com.exe test_report.pptx "给第1页标题添加动画淡入" --output test_out_cs_anim.pptx
            echo.
            echo --- 10e. C# 切换效果 ---
            pptx_editor_com.exe test_report.pptx "第1页切换效果淡化" --output test_out_cs_trans.pptx
            echo.
            echo --- 10f. C# 添加文本框 ---
            pptx_editor_com.exe test_report.pptx "添加文本框 内容是CSharp Test" --output test_out_cs_textbox.pptx
            echo.
            echo --- 10g. C# 居中对齐 ---
            pptx_editor_com.exe test_report.pptx "第1页标题居中对齐" --output test_out_cs_align.pptx
            echo.
            echo --- 10h. C# 背景填充 ---
            pptx_editor_com.exe test_report.pptx "第1页标题背景改成蓝色" --output test_out_cs_fill.pptx
            echo.
        ) else (
            echo ⚠️ C# 编译失败
            %CSC% /nologo /out:pptx_editor_com.exe /reference:"%PIA_PP%" /reference:"%PIA_OFFICE%" pptx_editor_com.cs
        )
    ) else (
        echo ⚠️ 未找到 Office Interop DLL，跳过 C# 测试
        echo    需要安装 Office PIA
    )
) else (
    echo ⚠️ 未找到 csc.exe，跳过 C# 测试
    echo    需要 .NET Framework 4.x
)
echo.

:done
echo.
echo ============================================
echo   ✅ 测试完成！
echo ============================================
echo.
echo 生成的文件:
dir /b *.pptx *.pdf *.png 2>nul
echo.

echo.
echo 💡 清理测试输出: demo_test.bat clean
pause
exit /b 0

:clean
echo 🧹 清理测试输出...
set count=0
for %%f in (*_modified.pptx test_out_*.pptx) do (if exist "%%f" (del "%%f" & echo   删除 %%f & set /a count+=1))
for %%f in (*_structure.json) do (if exist "%%f" (del "%%f" & echo   删除 %%f & set /a count+=1))
for %%f in (test_report.pptx) do (if exist "%%f" (del "%%f" & echo   删除 %%f & set /a count+=1))
for %%f in (test_report.pdf) do (if exist "%%f" (del "%%f" & echo   删除 %%f & set /a count+=1))
for %%f in (slide_*.png test_img.png) do (if exist "%%f" (del "%%f" & echo   删除 %%f & set /a count+=1))
for %%f in (pptx_editor_com.exe pptx_editor_com.pdb) do (if exist "%%f" (del "%%f" & echo   删除 %%f & set /a count+=1))
echo.
echo ✅ 清理完成
pause
