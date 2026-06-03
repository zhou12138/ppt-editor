@echo off
chcp 65001 >nul 2>&1
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
python pptx_editor.py test_report.pptx "把标题改成「Demo 测试成功」"
echo.

echo ========== 4. 测试：修改字号 ==========
python pptx_editor.py test_report_modified.pptx "第1页标题字号改成40"
echo.

echo ========== 5. 测试：加粗 ==========
python pptx_editor.py test_report_modified.pptx "第1页标题加粗"
echo.

echo ========== 6. 测试：改颜色 ==========
python pptx_editor.py test_report_modified.pptx "第1页标题颜色改成红色"
echo.

echo ========== 7. 测试：删除元素 ==========
python pptx_editor.py test_report_modified.pptx "删除第2页的表格"
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
python pptx_editor_com.py test_report.pptx "给第1页标题添加动画淡入"
echo.

echo --- 8c. COM 切换效果 ---
python pptx_editor_com.py test_report_modified.pptx "第1页切换效果淡化"
echo.

echo --- 8d. COM 导出 PDF ---
python pptx_editor_com.py test_report.pptx "导出PDF"
echo.

echo --- 8e. COM 导出图片 ---
python pptx_editor_com.py test_report.pptx --export-images
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
pause
