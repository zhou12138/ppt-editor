"""
PPT Editor COM - Direct API Test Suite
Opens PowerPoint ONCE and runs all tests sequentially.

Usage: python demo_test.py [--headed]
"""
import sys, os, traceback

headed = "--headed" in sys.argv

from pptx_editor_com import PowerPointCOM

passed = 0
failed = 0
errors = []

def test(name, fn):
    global passed, failed
    print(f"  === {name} ===")
    try:
        result = fn()
        if result is not None:
            print(f"    -> {result}")
        passed += 1
        print(f"    [OK]")
    except Exception as e:
        failed += 1
        errors.append((name, str(e)))
        print(f"    [FAIL] {e}")

def main():
    global passed, failed

    print("=" * 50)
    print("  PPT Editor - Direct API Test Suite")
    print(f"  Mode: {'HEADED' if headed else 'HEADLESS'}")
    print("=" * 50)

    print("\n--- Step 0: Generate test PPTX ---")
    import gen_test
    print("[OK] test_report.pptx generated\n")

    p = PowerPointCOM(visible=headed)
    try:
        p.open("test_report.pptx")
        # test_report.pptx has 4 slides:
        # 1: title + subtitle
        # 2: title + body (bullets)
        # 3: title + table 3x3
        # 4: 3 textboxes (blank layout)

        # Helper: find title on slide 1
        def s1t():
            return p.find_shape(1, {"type": "title"})[0]
        def s2b():
            return p.find_shape(2, {"type": "body"})[0]
        def s3tbl():
            return p.find_shape(3, {"type": "table"})[0]

        # ---- BASIC OPS ----
        print("\n---- BASIC OPS ----")
        test("inspect", lambda: f"Slides: {len(p.inspect()['slides'])}")

        # ---- TEXT OPS (on slide 1 title) ----
        print("\n---- TEXT OPS ----")
        test("modify_text", lambda: p.modify_text(s1t(), "New Title"))
        test("modify_partial_text", lambda: p.modify_partial_text(s1t(), 1, 3, "XX"))
        test("set_alignment_center", lambda: p.set_alignment(s1t(), "center"))

        # ---- FONT OPS ----
        print("\n---- FONT OPS ----")
        test("font_size_40", lambda: p.modify_font(s1t(), font_size=40))
        test("font_bold", lambda: p.modify_font(s1t(), bold=True))
        test("font_italic", lambda: p.modify_font(s1t(), italic=True))
        test("font_underline", lambda: p.modify_font(s1t(), underline=True))
        test("font_strikethrough", lambda: p.modify_font(s1t(), strikethrough=True))
        test("font_color_red", lambda: p.modify_font(s1t(), color=0x0000FF))
        test("font_name_Arial", lambda: p.modify_font(s1t(), font_name="Arial"))
        test("font_size_factor", lambda: p.modify_font(s1t(), font_size_factor=1.5))

        # ---- SHAPE OPS ----
        print("\n---- SHAPE OPS ----")
        test("set_fill_blue", lambda: p.set_fill(s1t(), 0xFF0000))
        test("set_border_red", lambda: p.set_border(s1t(), color_bgr=0x0000FF, weight=3))
        test("move_shape", lambda: p.move_shape(s1t(), left=50, top=50))
        test("resize_shape", lambda: p.resize_shape(s1t(), width=500, height=100))
        test("rotate_shape_45", lambda: p.rotate_shape(s1t(), 45))
        test("rotate_shape_back", lambda: p.rotate_shape(s1t(), -45))
        test("flip_horizontal", lambda: p.flip_shape(s1t(), "horizontal"))
        test("set_zorder_front", lambda: p.set_zorder(s1t(), "front"))

        # ---- TEXTBOX / PICTURE ----
        print("\n---- TEXTBOX / PICTURE ----")
        test("add_textbox", lambda: p.add_textbox(1, "Hello World", 100, 400, 300, 50))
        test("add_picture", lambda: p.add_picture(1, "test_img.png", 400, 400, 100, 75))

        # ---- ADVANCED TEXT ----
        print("\n---- ADVANCED TEXT ----")
        test("add_hyperlink", lambda: p.add_hyperlink(s1t(), "https://example.com"))
        test("set_line_spacing", lambda: p.set_line_spacing(s2b(), 1.5))
        test("set_paragraph_spacing", lambda: p.set_paragraph_spacing(s2b(), before=6, after=6))
        test("set_text_autofit", lambda: p.set_text_autofit(s1t(), "fit"))
        test("add_bullet", lambda: p.add_bullet(s2b(), 2))

        # ---- ADVANCED PICTURE ----
        print("\n---- ADVANCED PICTURE ----")
        def _find_pic():
            pics = p.find_shape(1, {"type": "picture"})
            if not pics:
                raise Exception("no picture on slide 1")
            return pics[0]
        test("crop_picture", lambda: p.crop_picture(_find_pic(), left=5, top=5, right=5, bottom=5))
        test("set_brightness", lambda: p.set_brightness(_find_pic(), 0.2))
        test("set_contrast", lambda: p.set_contrast(_find_pic(), 0.3))

        # ---- TABLE OPS (slide 3) ----
        print("\n---- TABLE OPS ----")
        test("modify_cell", lambda: p.modify_cell(3, {"type": "table"}, 1, 1, "Updated"))
        test("add_table_row", lambda: p.add_table_row(s3tbl()))
        test("add_table_column", lambda: p.add_table_column(s3tbl()))
        test("delete_table_row_last", lambda: p.delete_table_row(s3tbl(), s3tbl().Table.Rows.Count))
        test("delete_table_column_last", lambda: p.delete_table_column(s3tbl(), s3tbl().Table.Columns.Count))
        test("add_table_on_slide1", lambda: p.add_table(1, 2, 2, 100, 200, 300, 100))

        # ---- ANIMATION / TRANSITION ----
        print("\n---- ANIMATION / TRANSITION ----")
        test("add_animation_fade", lambda: p.add_animation(1, s1t(), "fade"))
        test("modify_animation_effect", lambda: p.modify_animation_effect(1, 1, "zoom"))
        test("remove_animation", lambda: p.remove_animation(1))
        test("set_transition_fade", lambda: p.set_transition(1, "fade", 1.5))
        test("set_transition_wipe", lambda: p.set_transition(2, "wipe", 1.0))

        # ---- 3D / VISUAL EFFECTS ----
        print("\n---- 3D / VISUAL EFFECTS ----")
        test("set_shadow", lambda: p.set_shadow(s1t(), 1))
        test("set_3d_rotation", lambda: p.set_3d_rotation(s1t(), x=10, y=10, z=0))

        # ---- NOTES ----
        print("\n---- NOTES ----")
        test("set_notes", lambda: p.set_notes(1, "Speaker notes for slide 1"))
        test("get_notes", lambda: f"Notes: {repr(p.get_notes(1))}")

        # ---- COMMENTS ----
        print("\n---- COMMENTS ----")
        test("add_comment", lambda: p.add_comment(1, "Review this slide", "Tester", 10, 10))
        test("get_comments", lambda: f"Comments: {p.get_comments(1)}")
        test("delete_comment", lambda: p.delete_comment(1, 1))

        # ---- SECTIONS ----
        print("\n---- SECTIONS ----")
        test("add_section", lambda: p.add_section("Intro", 1))
        test("get_sections", lambda: f"Sections: {p.get_sections()}")
        test("rename_section", lambda: p.rename_section(1, "Introduction"))
        test("delete_section", lambda: p.delete_section(1))

        # ---- BACKGROUND ----
        print("\n---- BACKGROUND ----")
        test("set_slide_background_blue", lambda: p.set_slide_background(1, 0xFF0000))
        test("set_slide_background_image", lambda: p.set_slide_background_image(2, "test_img.png"))

        # ---- CHART ----
        print("\n---- CHART ----")
        # Use xlColumnClustered=51 with data
        test("add_chart_bar", lambda: p.add_chart(4, 51, data={
            "categories": ["A", "B", "C"],
            "series": [{"name": "S1", "values": [10, 20, 30]}]
        }))

        # ---- FREEFORM ----
        print("\n---- FREEFORM ----")
        test("add_freeform", lambda: p.add_freeform(1, [(100, 300), (200, 300), (200, 350), (100, 350)]))

        # ---- MASTER / LAYOUT ----
        print("\n---- MASTER / LAYOUT ----")
        test("get_slide_layouts", lambda: f"Layouts: {p.get_slide_layouts()}")
        test("get_slide_masters", lambda: f"Masters: {p.get_slide_masters()}")
        test("set_slide_layout", lambda: p.set_slide_layout(4, 1))

        # ---- THEME ----
        print("\n---- THEME ----")
        test("get_theme_colors", lambda: f"Theme colors: {p.get_theme_colors()}")

        # ---- SLIDESHOW SETTINGS ----
        print("\n---- SLIDESHOW SETTINGS ----")
        test("set_slideshow_settings", lambda: p.set_slideshow_settings(
            loop=False, show_narration=True, show_animation=True))

        # ---- EXPORT ----
        print("\n---- EXPORT ----")
        test("export_pdf", lambda: p.export_pdf("test_out.pdf"))
        test("export_image_slide1", lambda: p.export_image(1, "slide_1.png"))

        # ---- SAVE ----
        print("\n---- SAVE ----")
        test("save_as_test_out", lambda: (p.save("test_out.pptx"), "saved")[1])

        # ---- SLIDE OPS (do last, changes structure) ----
        print("\n---- SLIDE OPS (last, changes structure) ----")
        test("add_slide_blank", lambda: p.add_slide(layout=12))
        test("duplicate_slide_1", lambda: p.duplicate_slide(1))
        test("set_slide_size", lambda: p.set_slide_size(720, 540))
        cnt = p.prs.Slides.Count
        test("move_slide_last_to_1", lambda: p.move_slide(cnt, 1))
        test("delete_slide_last", lambda: p.delete_slide(p.prs.Slides.Count))
        # delete the other added slide
        test("delete_slide_last2", lambda: p.delete_slide(p.prs.Slides.Count))

        # ---- MERGE ----
        print("\n---- MERGE ----")
        test("merge_presentations", lambda: p.merge_presentations(
            ["test_report.pptx"], "test_merged.pptx"))

        # Final save
        test("save_final", lambda: (p.save("test_out_final.pptx"), "saved")[1])

    finally:
        p.close()

    # ---- SUMMARY ----
    print("\n" + "=" * 50)
    print(f"  RESULTS: {passed} passed, {failed} failed, {passed + failed} total")
    print("=" * 50)
    if errors:
        print("\nFailed tests:")
        for name, err in errors:
            print(f"  - {name}: {err}")
    print()
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
