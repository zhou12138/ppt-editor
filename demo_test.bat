@echo off
cd /d "%~dp0"
if "%1"=="clean" goto CLEAN

echo ============================================
echo   PPT Editor - Direct API Test Demo
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
echo === Step 0: Generate test PPTX ===
call python gen_test.py
if errorlevel 1 (
    echo [FAIL] gen_test.py failed
    pause
    exit /b 1
)
echo [OK] test_report.pptx generated
echo.
echo ---- BASIC OPS ----
echo === Test: open_inspect_close ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); d=p.inspect(); print('Slides:',len(d['slides'])); p.close()"
echo.
echo === Test: save_as ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); p.save('test_out.pptx'); p.close()"
echo.
echo ---- TEXT OPS ----
echo === Test: modify_text ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_text(s[0],'New Title')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: modify_partial_text ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_partial_text(s[0],1,2,'XX')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_alignment ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.set_alignment(s[0],'center')); p.save('test_out.pptx'); p.close()"
echo.
echo ---- FONT OPS ----
echo === Test: modify_font_size ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_font(s[0],font_size=40)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: modify_font_bold ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_font(s[0],bold=True)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: modify_font_italic ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_font(s[0],italic=True)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: modify_font_underline ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_font(s[0],underline=True)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: modify_font_color ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_font(s[0],color=0x0000FF)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: modify_font_name ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.modify_font(s[0],font_name='Arial')); p.save('test_out.pptx'); p.close()"
echo.
echo ---- SHAPE OPS ----
echo === Test: set_fill ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.set_fill(s[0],0xFF0000)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_border ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.set_border(s[0],color_bgr=0x0000FF,weight=3)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: move_shape ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.move_shape(s[0],left=50,top=50)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: resize_shape ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.resize_shape(s[0],width=500,height=100)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: delete_shape ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'subtitle'}); print(p.delete_shape(s[0])); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: rotate_shape ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.rotate_shape(s[0],45)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: flip_shape ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.flip_shape(s[0],'horizontal')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_zorder ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.set_zorder(s[0],'front')); p.save('test_out.pptx'); p.close()"
echo.
echo ---- TEXTBOX / PICTURE ----
echo === Test: add_textbox ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_textbox(1,'Hello World',100,100,300,50)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: add_picture ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_picture(1,'test_img.png',100,100,200,150)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- SLIDE OPS ----
echo === Test: add_slide ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_slide(layout=12)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: delete_slide ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print(p.delete_slide(5)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: move_slide ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.move_slide(3,1)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: duplicate_slide ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.duplicate_slide(1)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_slide_size ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.set_slide_size(720,540)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- TABLE OPS ----
echo === Test: add_table ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_table(1,3,3,100,200,400,200)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: modify_cell ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.modify_cell(3,{'type':'table'},1,1,'Updated')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: add_table_row ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(3,{'type':'table'}); print(p.add_table_row(s[0])); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: delete_table_row ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); s=p.find_shape(3,{'type':'table'}); print(p.delete_table_row(s[0],4)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: add_table_column ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(3,{'type':'table'}); print(p.add_table_column(s[0])); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: delete_table_column ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); s=p.find_shape(3,{'type':'table'}); print(p.delete_table_column(s[0],4)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- ANIMATION / TRANSITION ----
echo === Test: add_animation ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.add_animation(1,s[0],'fade')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: remove_animation ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print(p.remove_animation(1)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_transition ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.set_transition(1,'fade',1.5)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- EXPORT ----
echo === Test: export_pdf ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.export_pdf('test_out.pdf')); p.close()"
echo.
echo === Test: export_image ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.export_image(1,'slide_1.png')); p.close()"
echo.
echo ---- NOTES ----
echo === Test: set_notes ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.set_notes(1,'Speaker notes here')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: get_notes ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print('Notes:',repr(p.get_notes(1))); p.close()"
echo.
echo ---- COMMENTS ----
echo === Test: add_comment ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_comment(1,'Test comment','Tester',10,10)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: get_comments ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print('Comments:',p.get_comments(1)); p.close()"
echo.
echo === Test: delete_comment ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print(p.delete_comment(1,1)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- SECTIONS ----
echo === Test: add_section ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_section('Intro',1)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: get_sections ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print('Sections:',p.get_sections()); p.close()"
echo.
echo === Test: rename_section ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print(p.rename_section(1,'Introduction')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: delete_section ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); print(p.delete_section(1)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- BACKGROUND ----
echo === Test: set_slide_background ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.set_slide_background(1,0xFF0000)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_slide_background_image ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.set_slide_background_image(1,'test_img.png')); p.save('test_out.pptx'); p.close()"
echo.
echo ---- ADVANCED TEXT ----
echo === Test: add_hyperlink ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.add_hyperlink(s[0],'https://example.com')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_line_spacing ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(2,{'type':'body'}); print(p.set_line_spacing(s[0],1.5)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_paragraph_spacing ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(2,{'type':'body'}); print(p.set_paragraph_spacing(s[0],before=6,after=6)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_text_autofit ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.set_text_autofit(s[0],'fit')); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: add_bullet ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(2,{'type':'body'}); print(p.add_bullet(s[0],2)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- ADVANCED PICTURE ----
echo === Test: crop_picture ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); p.add_picture(1,'test_img.png',100,100,200,150); s=p.find_shape(1,{'type':'picture'}); print(p.crop_picture(s[0],left=10,top=10,right=10,bottom=10)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_brightness ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); s=p.find_shape(1,{'type':'picture'}); print(p.set_brightness(s[0],0.3)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_contrast ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); s=p.find_shape(1,{'type':'picture'}); print(p.set_contrast(s[0],0.5)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: replace_picture ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_out.pptx'); s=p.find_shape(1,{'type':'picture'}); print(p.replace_picture(s[0],'test_img.png')); p.save('test_out.pptx'); p.close()"
echo.
echo ---- CHART ----
echo === Test: add_chart ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_chart(1,51)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- 3D / VISUAL EFFECTS ----
echo === Test: set_shadow ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.set_shadow(s[0],1)); p.save('test_out.pptx'); p.close()"
echo.
echo === Test: set_3d_rotation ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); s=p.find_shape(1,{'type':'title'}); print(p.set_3d_rotation(s[0],x=10,y=10,z=0)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- FREEFORM ----
echo === Test: add_freeform ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.add_freeform(1,[(100,100),(200,100),(200,200),(100,200)])); p.save('test_out.pptx'); p.close()"
echo.
echo ---- MASTER / LAYOUT ----
echo === Test: get_slide_layouts ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print('Layouts:',p.get_slide_layouts()); p.close()"
echo.
echo === Test: get_slide_masters ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print('Masters:',p.get_slide_masters()); p.close()"
echo.
echo === Test: set_slide_layout ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.set_slide_layout(1,1)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- THEME ----
echo === Test: get_theme_colors ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print('Theme colors:',p.get_theme_colors()); p.close()"
echo.
echo ---- SLIDESHOW SETTINGS ----
echo === Test: set_slideshow_settings ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.set_slideshow_settings(loop=False,show_narration=True,show_animation=True)); p.save('test_out.pptx'); p.close()"
echo.
echo ---- MERGE ----
echo === Test: merge_presentations ===
call python -c "from pptx_editor_com import PowerPointCOM; p=PowerPointCOM(); p.open('test_report.pptx'); print(p.merge_presentations(['test_report.pptx'],'test_merged.pptx')); p.close()"
echo.

echo.
echo ============================================
echo   All tests completed!
echo ============================================
echo.
echo Generated files:
dir /b test_out*.pptx test_out*.pdf test_merged*.pptx slide_*.png 2>nul
echo.
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
