"""
PPT Editor COM - Demo Test Suite
Builds a polished presentation while exercising all COM APIs.
Each test contributes to a visually coherent final PPTX.

Usage: python demo_test.py [--headed]
"""
import sys, os, time

# Always work from the script's own directory
os.chdir(os.path.dirname(os.path.abspath(__file__)))

headed = "--headed" in sys.argv
DELAY = 1.5 if headed else 0

from pptx_editor_com import PowerPointCOM

passed = 0
failed = 0
errors = []
current_slide = [0]  # mutable for closure
slide_notes = {}  # {slide_num: [lines]}
ppt_ref = [None]  # ref to PowerPointCOM for notes writing

def test(name, fn):
    global passed, failed
    print(f"  {name}", end=" ... ")
    try:
        result = fn()
        passed += 1
        detail = ""
        msg = f"OK"
        if result and isinstance(result, str) and len(result) < 80:
            msg += f" ({result})"
            detail = result
        print(msg)
        _write_note(name, "✅", detail)
        if DELAY > 0:
            time.sleep(DELAY)
    except Exception as e:
        failed += 1
        errors.append((name, str(e)))
        print(f"FAIL: {e}")
        _write_note(name, "❌", str(e)[:60])

def _write_note(name, status, detail):
    sn = current_slide[0]
    p = ppt_ref[0]
    if sn <= 0 or p is None:
        return
    line = f"{status} {name}"
    if detail:
        line += f" → {detail}"
    slide_notes.setdefault(sn, []).append(line)
    try:
        text = "\n".join(slide_notes[sn])
        p.set_notes(sn, text)
    except:
        pass

def main():
    global passed, failed

    print("=" * 55)
    print("  PPT Editor COM - Demo Test Suite")
    print(f"  Mode: {'HEADED (visual)' if headed else 'HEADLESS (fast)'}")
    print("=" * 55)

    print("\n[Step 0] Generating base PPTX...")
    import gen_test
    print()

    p = PowerPointCOM(visible=headed)
    try:
        p.open("test_report.pptx")
        ppt_ref[0] = p

        # Helpers
        def goto(n):
            if headed:
                try: p.app.ActiveWindow.View.GotoSlide(n)
                except: pass

        def title(n):
            return p.find_shape(n, {"type": "title"})[0]

        def body(n):
            return p.find_shape(n, {"type": "body"})[0]

        def tbl(n):
            return p.find_shape(n, {"type": "table"})[0]

        def textboxes(n):
            """All non-placeholder textbox shapes on slide n."""
            shapes = []
            for i in range(1, p.prs.Slides(n).Shapes.Count + 1):
                s = p.prs.Slides(n).Shapes(i)
                if s.HasTextFrame and s.Type == 17:  # msoTextBox
                    shapes.append(s)
            return shapes

        # ============================================================
        # SLIDE 1: Title Page - Professional styling
        # ============================================================
        print("\n--- Slide 1: Title Page ---")
        current_slide[0] = 1
        goto(1)

        # Inspect
        test("inspect", lambda: f"{len(p.inspect()['slides'])} slides")

        # Style the title
        test("set_title", lambda: p.modify_text(title(1), "PPT Editor"))
        test("title_font", lambda: p.modify_font(title(1), font_size=54, bold=True, font_name="Calibri"))
        test("title_color", lambda: p.modify_font(title(1), color=0x4A3A2D))  # dark blue-gray BGR
        test("title_align", lambda: p.set_alignment(title(1), "center"))

        # Style subtitle
        def sub(n):
            return p.find_shape(n, {"type": "subtitle"})[0]
        test("set_subtitle", lambda: p.modify_text(sub(1), "COM Automation PowerShow"))
        test("subtitle_font", lambda: p.modify_font(sub(1), font_size=24, italic=True))
        test("subtitle_color", lambda: p.modify_font(sub(1), color=0x8B7B6B))  # warm gray

        # Background - gradient dark blue
        test("slide1_bg", lambda: p.set_slide_background(1, 0x4A3A2D))  # dark navy

        # Add a decorative textbox
        test("add_tagline", lambda: p.add_textbox(1, "70+ COM Methods | Real-time Visual Demo", 200, 480, 560, 40))

        # Add company logo placeholder (picture)
        test("add_logo", lambda: p.add_picture(1, "test_img.png", 580, 40, 80, 60))

        # Notes
        test("slide1_notes", lambda: p.set_notes(1, "Title slide - introduces PPT Editor COM capabilities"))

        # Transition
        test("slide1_transition", lambda: p.set_transition(1, "fade", 1.0))

        # ============================================================
        # SLIDE 2: Feature List - Enhanced typography
        # ============================================================
        print("\n--- Slide 2: Feature Overview ---")
        current_slide[0] = 2
        goto(2)

        test("s2_title", lambda: p.modify_text(title(2), "What Can It Do?"))
        test("s2_title_font", lambda: p.modify_font(title(2), font_size=40, bold=True, color=0x783C14))
        test("s2_title_align", lambda: p.set_alignment(title(2), "left"))

        # Enhance the body text
        test("s2_body_font", lambda: p.modify_font(body(2), font_size=20, font_name="Calibri"))
        test("s2_line_spacing", lambda: p.set_line_spacing(body(2), 1.8))
        test("s2_para_spacing", lambda: p.set_paragraph_spacing(body(2), before=6, after=4))
        test("s2_bullets", lambda: p.add_bullet(body(2), 1))

        # Background
        test("s2_bg", lambda: p.set_slide_background(2, 0xFAF5F0))  # warm off-white

        # Add a side accent box
        test("s2_accent", lambda: p.add_textbox(2, "", 20, 120, 12, 380))
        # Fill the accent bar
        def _fill_accent():
            boxes = textboxes(2)
            for b in boxes:
                if b.Width < 20:  # the thin accent bar
                    p.set_fill(b, 0x783C14)  # orange accent
                    return "accent bar filled"
            return "no accent bar found"
        test("s2_accent_fill", _fill_accent)

        test("s2_transition", lambda: p.set_transition(2, "wipe", 0.8))
        test("s2_notes", lambda: p.set_notes(2, "Feature overview - lists all COM capabilities"))

        # ============================================================
        # SLIDE 3: Data Table - Professional formatting
        # ============================================================
        print("\n--- Slide 3: Performance Table ---")
        current_slide[0] = 3
        goto(3)

        test("s3_title", lambda: p.modify_text(title(3), "Q2 Performance Dashboard"))
        test("s3_title_font", lambda: p.modify_font(title(3), font_size=36, bold=True, color=0x783C14))

        # Modify table cells with real data emphasis
        test("s3_cell_header", lambda: p.modify_cell(3, {"type": "table"}, 1, 1, "KPI"))
        test("s3_cell_q1", lambda: p.modify_cell(3, {"type": "table"}, 1, 2, "Q1 2026"))
        test("s3_cell_q2", lambda: p.modify_cell(3, {"type": "table"}, 1, 3, "Q2 2026"))
        test("s3_cell_growth", lambda: p.modify_cell(3, {"type": "table"}, 1, 4, "YoY Growth"))

        # Add a row for new metric
        test("s3_add_row", lambda: p.add_table_row(tbl(3)))
        # Fill new row
        test("s3_new_metric", lambda: p.modify_cell(3, {"type": "table"}, 6, 1, "Retention"))
        test("s3_new_q1", lambda: p.modify_cell(3, {"type": "table"}, 6, 2, "72%"))
        test("s3_new_q2", lambda: p.modify_cell(3, {"type": "table"}, 6, 3, "85%"))
        test("s3_new_growth", lambda: p.modify_cell(3, {"type": "table"}, 6, 4, "+18%"))

        # Add column then remove it (demo the capability)
        test("s3_add_col", lambda: p.add_table_column(tbl(3)))
        test("s3_del_col", lambda: p.delete_table_column(tbl(3), tbl(3).Table.Columns.Count))

        # Background
        test("s3_bg", lambda: p.set_slide_background(3, 0xFFF8F2))
        test("s3_transition", lambda: p.set_transition(3, "fade", 1.0))

        # ============================================================
        # SLIDE 4: Chart & Shapes - Data visualization
        # ============================================================
        print("\n--- Slide 4: Data Visualization ---")
        current_slide[0] = 4
        goto(4)

        # Add chart
        test("s4_chart", lambda: p.add_chart(4, 51, data={
            "categories": ["Jan", "Feb", "Mar", "Apr", "May", "Jun"],
            "series": [
                {"name": "Revenue", "values": [120, 135, 148, 162, 175, 195]},
                {"name": "Target",  "values": [130, 140, 150, 160, 170, 180]},
            ]
        }, left=40, top=120, width=550, height=380))

        # Add a callout textbox
        test("s4_callout", lambda: p.add_textbox(4, "Revenue exceeded target in Q2!", 420, 80, 250, 40))
        def _style_callout():
            boxes = textboxes(4)
            for b in boxes:
                if "exceeded" in (b.TextFrame.TextRange.Text if b.HasTextFrame else ""):
                    p.modify_font(b, font_size=14, bold=True, color=0x008B00)  # green
                    p.set_fill(b, 0xE8FFE8)  # light green bg
                    p.set_border(b, color_bgr=0x008B00, weight=1.5)
                    return "callout styled"
            return "no callout found"
        test("s4_callout_style", _style_callout)

        # Add picture
        test("s4_picture", lambda: p.add_picture(4, "test_img.png", 620, 420, 100, 75))

        # Freeform decorative shape
        test("s4_freeform", lambda: p.add_freeform(4, [
            (620, 130), (700, 130), (720, 200), (700, 270), (620, 270), (600, 200)
        ]))

        test("s4_bg", lambda: p.set_slide_background(4, 0xFFF8F2))
        test("s4_transition", lambda: p.set_transition(4, "dissolve", 1.0))

        # ============================================================
        # SLIDE 5: Visual Effects Showcase
        # ============================================================
        print("\n--- Slide 5: Visual Effects ---")
        current_slide[0] = 5
        goto(5)

        # Add styled boxes to demonstrate effects
        test("s5_box1", lambda: p.add_textbox(5, "Shadow Effect", 80, 140, 250, 80))
        test("s5_box2", lambda: p.add_textbox(5, "3D Rotation", 380, 140, 250, 80))
        test("s5_box3", lambda: p.add_textbox(5, "Styled Border", 80, 300, 250, 80))
        test("s5_box4", lambda: p.add_textbox(5, "Color Fill", 380, 300, 250, 80))

        def _style_effect_boxes():
            boxes = textboxes(5)
            results = []
            for b in boxes:
                txt = b.TextFrame.TextRange.Text if b.HasTextFrame else ""
                p.modify_font(b, font_size=18, bold=True, font_name="Calibri")

                if "Shadow" in txt:
                    p.set_fill(b, 0xFFFFFF)
                    p.modify_font(b, color=0x4A3A2D)
                    p.set_shadow(b, 1)
                    results.append("shadow")

                elif "3D" in txt:
                    p.set_fill(b, 0xE8D8C8)
                    p.modify_font(b, color=0x4A3A2D)
                    p.set_3d_rotation(b, x=8, y=5, z=0)
                    results.append("3d")

                elif "Border" in txt:
                    p.set_border(b, color_bgr=0x783C14, weight=3)
                    p.modify_font(b, color=0x783C14)
                    results.append("border")

                elif "Fill" in txt:
                    p.set_fill(b, 0x783C14)  # orange
                    p.modify_font(b, color=0xFFFFFF)
                    results.append("fill")

            return ", ".join(results)

        test("s5_style_all", _style_effect_boxes)

        # Add an animated textbox
        test("s5_anim_box", lambda: p.add_textbox(5, "Animated!", 280, 440, 150, 50))
        def _animate_box():
            boxes = textboxes(5)
            for b in boxes:
                if "Animated" in (b.TextFrame.TextRange.Text if b.HasTextFrame else ""):
                    p.modify_font(b, font_size=20, bold=True, color=0x0000CC)
                    p.set_fill(b, 0xE8E8FF)
                    p.add_animation(5, b, "fade")
                    return "animated"
            return "no box"
        test("s5_animate", _animate_box)

        # Background image
        test("s5_bg", lambda: p.set_slide_background(5, 0xFAF5F0))
        test("s5_transition", lambda: p.set_transition(5, "fade", 1.2))

        # ============================================================
        # SLIDE 6: Thank You - clean closing
        # ============================================================
        print("\n--- Slide 6: Thank You ---")
        current_slide[0] = 6
        goto(6)

        test("s6_title", lambda: p.modify_text(title(6), "Thank You!"))
        test("s6_title_font", lambda: p.modify_font(title(6), font_size=52, bold=True, color=0x4A3A2D))
        test("s6_title_align", lambda: p.set_alignment(title(6), "center"))

        test("s6_subtitle", lambda: p.modify_text(sub(6), "Built with PPT Editor COM\n70+ methods | Full automation"))
        test("s6_subtitle_font", lambda: p.modify_font(sub(6), font_size=22, color=0x8B7B6B))

        test("s6_bg", lambda: p.set_slide_background(6, 0x4A3A2D))
        test("s6_transition", lambda: p.set_transition(6, "fade", 1.5))
        test("s6_notes", lambda: p.set_notes(6, "Closing slide - thank the audience"))

        # ============================================================
        # CROSS-SLIDE OPS
        # ============================================================
        print("\n--- Cross-slide Operations ---")
        current_slide[0] = 0

        # Comments
        test("add_comment", lambda: p.add_comment(1, "Great title design!", "Reviewer", 100, 50))
        test("get_comments", lambda: f"{len(p.get_comments(1))} comments")
        test("del_comment", lambda: p.delete_comment(1, 1))

        # Sections
        test("add_section1", lambda: p.add_section("Introduction", 1))
        test("add_section2", lambda: p.add_section("Content", 2))
        test("get_sections", lambda: f"{len(p.get_sections())} sections")
        test("rename_section", lambda: p.rename_section(1, "Opening"))
        test("del_section2", lambda: p.delete_section(2))
        test("del_section1", lambda: p.delete_section(1))

        # Layouts & theme info
        test("get_layouts", lambda: f"{len(p.get_slide_layouts())} layouts")
        test("get_masters", lambda: f"{len(p.get_slide_masters())} masters")
        test("get_theme", lambda: f"{len(p.get_theme_colors())} theme colors")

        # Slideshow settings
        test("slideshow_cfg", lambda: p.set_slideshow_settings(
            loop=False, show_narration=True, show_animation=True))

        # Hyperlink on title
        goto(1)
        test("add_hyperlink", lambda: p.add_hyperlink(title(1), "https://github.com"))

        # Text autofit
        test("autofit", lambda: p.set_text_autofit(title(1), "fit"))

        # ============================================================
        # EXPORT & SAVE
        # ============================================================
        print("\n--- Export & Save ---")
        goto(1)
        test("export_pdf", lambda: p.export_pdf("demo_output.pdf"))
        test("export_slide1_png", lambda: p.export_image(1, "demo_slide1.png"))
        test("save_pptx", lambda: (p.save("demo_output.pptx"), "saved")[1])

        # ============================================================
        # SLIDE STRUCTURE OPS (last - changes indices)
        # ============================================================
        print("\n--- Slide Structure Ops ---")
        test("add_blank_slide", lambda: p.add_slide(layout=7))
        goto(p.prs.Slides.Count)
        test("duplicate_slide1", lambda: p.duplicate_slide(1))

        cnt = p.prs.Slides.Count
        test("move_slide", lambda: p.move_slide(cnt, cnt - 1))
        test("del_extra1", lambda: p.delete_slide(p.prs.Slides.Count))
        test("del_extra2", lambda: p.delete_slide(p.prs.Slides.Count))
        test("set_slide_size", lambda: p.set_slide_size(960, 540))  # 16:9

        # Merge
        test("merge", lambda: p.merge_presentations(["test_report.pptx"], "demo_merged.pptx"))

        # Final save
        goto(1)
        test("final_save", lambda: (p.save("demo_final.pptx"), "saved")[1])

    finally:
        if headed:
            # Flip through all slides as a review
            print(f"\n  [HEADED] Reviewing slides...")
            for i in range(1, p.prs.Slides.Count + 1):
                goto(i)
                time.sleep(4)
            print(f"  [HEADED] Closing in 10s...")
            time.sleep(10)
        p.close()

    # ---- SUMMARY ----
    print("\n" + "=" * 55)
    print(f"  RESULTS: {passed} passed, {failed} failed, {passed + failed} total")
    if failed == 0:
        print("  All tests passed! Check demo_final.pptx")
    print("=" * 55)
    if errors:
        print("\nFailed tests:")
        for name, err in errors:
            print(f"  - {name}: {err}")
    print()
    return 0 if failed == 0 else 1

if __name__ == "__main__":
    sys.exit(main())
