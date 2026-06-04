"""
PPTX 自然语言编辑器 — COM 版本 (Windows Only)
需要: pip install pywin32 + Microsoft Office

用法: python pptx_editor_com.py <pptx文件> [指令]
      python pptx_editor_com.py <pptx文件> --inspect
      python pptx_editor_com.py <pptx文件> --interactive

COM 独有能力（python-pptx 做不到的）:
  - 添加动画: "给第1页标题添加动画淡入"
  - 切换效果: "第2页切换效果淡化"
  - 导出PDF:  "导出PDF"
  - 导出图片: --export-images
  - 删除/修改动画: "删除动画" / "清除动画"
  - 添加文本框/图片/幻灯片管理/表格操作/对齐/填充/边框/移动/缩放
"""
import sys, os, re, json

try:
    import win32com.client
    import pythoncom
except ImportError:
    print("❌ pip install pywin32 (Windows + Office required)")
    sys.exit(1)

# ========== 颜色 (COM 用 BGR!) ==========
COLOR_MAP = {
    "红": 0x0000FF, "红色": 0x0000FF, "蓝": 0xFF0000, "蓝色": 0xFF0000,
    "绿": 0x00AA00, "绿色": 0x00AA00, "黄": 0x00D7FF, "黄色": 0x00D7FF,
    "黑": 0x000000, "黑色": 0x000000, "白": 0xFFFFFF, "白色": 0xFFFFFF,
    "灰": 0x888888, "灰色": 0x888888, "橙": 0x008CFF, "橙色": 0x008CFF,
    "紫": 0x800080, "紫色": 0x800080, "粉": 0xB469FF, "粉色": 0xB469FF,
    "red": 0x0000FF, "blue": 0xFF0000, "green": 0x00AA00,
    "black": 0x000000, "white": 0xFFFFFF,
}

PH_NAMES = {1:"TITLE",2:"BODY",3:"CENTER_TITLE",4:"SUBTITLE",7:"OBJECT",
            8:"CHART",9:"TABLE",12:"MEDIA",13:"SLIDE_NUMBER",15:"FOOTER"}

class PowerPointCOM:
    def __init__(self, visible=False):
        pythoncom.CoInitialize()
        self.app = win32com.client.Dispatch("PowerPoint.Application")
        if visible: self.app.Visible = True
        self.prs = None

    def _detect_shape_content(self, shape):
        info = {"has_image": False, "has_chart": False, "has_table": False, "has_media": False}

        try:
            info["has_table"] = bool(shape.HasTable)
        except Exception:
            pass

        try:
            info["has_chart"] = bool(shape.HasChart)
        except Exception:
            pass

        try:
            if shape.Type == 13:
                info["has_image"] = True
            elif shape.Type == 21:
                info["has_media"] = True
        except Exception:
            pass

        try:
            shape.PictureFormat
            info["has_image"] = True
        except Exception:
            pass

        try:
            contained_type = shape.PlaceholderFormat.ContainedType
            if contained_type == 13:
                info["has_image"] = True
            elif contained_type == 8:
                info["has_chart"] = True
            elif contained_type == 19:
                info["has_table"] = True
            elif contained_type == 21:
                info["has_media"] = True
        except Exception:
            pass

        return info

    def _format_shape_summary(self, element):
        if element.get("text"):
            return element["text"][:40].replace("\n", "↵")

        labels = []
        if element.get("has_image"):
            labels.append("[图片]")
        if element.get("has_chart"):
            labels.append("[图表]")
        if element.get("table"):
            labels.append(f"[表格 {len(element['table'])}×{len(element['table'][0])}]")
        elif element.get("has_table"):
            labels.append("[表格]")
        if element.get("has_media"):
            labels.append("[媒体]")
        return " ".join(labels) if labels else "(无)"

    def _get_save_format(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext == ".pptx":
            return 24  # ppSaveAsOpenXMLPresentation
        if ext == ".ppt":
            return 1  # ppSaveAsPresentation
        return None

    def _save_presentation(self, path):
        fmt = self._get_save_format(path)
        if fmt is None:
            self.prs.SaveAs(path)
        else:
            self.prs.SaveAs(path, fmt)

    def open(self, path):
        self.filepath = os.path.abspath(path)
        with_window = self.app.Visible
        self.prs = self.app.Presentations.Open(self.filepath, False, False, with_window)
        print(f"📂 已打开: {self.filepath} ({self.prs.Slides.Count}页)")
        return self

    def close(self):
        try:
            if self.prs: self.prs.Close()
            if self.app: self.app.Quit()
        except: pass
        pythoncom.CoUninitialize()

    def save(self, out=None):
        path = os.path.abspath(out) if out else self.filepath
        self._save_presentation(path)
        if out:
            print(f"💾 已另存为: {out}")
        else:
            print(f"💾 已保存")

    # ---- 结构查看 ----
    def inspect(self):
        result = {"slides": []}
        sw = self.prs.PageSetup.SlideWidth
        sh = self.prs.PageSetup.SlideHeight
        for si in range(1, self.prs.Slides.Count + 1):
            slide = self.prs.Slides(si)
            sd = {"index": si, "layout": "", "elements": []}
            try: sd["layout"] = slide.CustomLayout.Name
            except: sd["layout"] = str(slide.Layout)
            
            for shape in slide.Shapes:
                e = {"id": shape.Id, "name": shape.Name, "type": shape.Type,
                     "left": round(shape.Left,1), "top": round(shape.Top,1),
                     "width": round(shape.Width,1), "height": round(shape.Height,1),
                     "text": "", "is_placeholder": False,
                     "has_image": False, "has_chart": False,
                     "has_table": False, "has_media": False}
                
                cx = shape.Left + shape.Width/2
                cy = shape.Top + shape.Height/2
                h = "左" if cx < sw*0.33 else ("右" if cx > sw*0.67 else "中")
                v = "上" if cy < sh*0.33 else ("下" if cy > sh*0.67 else "中")
                e["position_label"] = h + v
                
                try:
                    pf = shape.PlaceholderFormat
                    if pf:
                        e["is_placeholder"] = True
                        e["ph_type"] = pf.Type
                        e["ph_type_name"] = PH_NAMES.get(pf.Type, f"({pf.Type})")
                except: pass

                e.update(self._detect_shape_content(shape))
                
                try:
                    if shape.HasTextFrame:
                        e["text"] = shape.TextFrame.TextRange.Text
                        e["paragraphs"] = []
                        for pi in range(1, shape.TextFrame.TextRange.Paragraphs().Count + 1):
                            p = shape.TextFrame.TextRange.Paragraphs(pi)
                            pd = {"text": p.Text, "font": p.Font.Name,
                                  "size": p.Font.Size, "bold": bool(p.Font.Bold)}
                            try: pd["color"] = p.Font.Color.RGB
                            except: pass
                            e["paragraphs"].append(pd)
                except: pass
                
                try:
                    if shape.HasTable:
                        t = shape.Table
                        e["table"] = [[t.Cell(r,c).Shape.TextFrame.TextRange.Text
                                       for c in range(1,t.Columns.Count+1)]
                                      for r in range(1,t.Rows.Count+1)]
                except: pass
                
                sd["elements"].append(e)
            result["slides"].append(sd)
        return result

    def print_structure(self, desc):
        for s in desc["slides"]:
            print(f"\n{'='*50}\n📄 第 {s['index']} 页 ({s['layout']})\n{'='*50}")
            for e in s["elements"]:
                ph = f" [{e.get('ph_type_name','')}]" if e.get("is_placeholder") else ""
                txt = self._format_shape_summary(e)
                idx = s["elements"].index(e) + 1
                print(f"  [{idx}] [{e['id']}] {e['name']}{ph} ({e['position_label']}) → {txt}")
                if e.get("paragraphs"):
                    p = e["paragraphs"][0]
                    print(f"       字体:{p.get('font')} 字号:{p.get('size')} 粗:{p.get('bold')}")
                if e.get("table"):
                    print(f"       表格: {len(e['table'])}×{len(e['table'][0])}")

    # ---- 元素定位 ----
    def find_shape(self, slide_idx, target):
        slide = self.prs.Slides(slide_idx)
        hits = []
        for shape_index, shape in enumerate(slide.Shapes, 1):
            ok = True
            if "type" in target:
                t = target["type"]
                try:
                    pf = shape.PlaceholderFormat
                    if t == "title" and pf.Type not in (1,3): ok = False
                    elif t == "subtitle" and pf.Type != 4: ok = False
                    elif t == "body" and pf.Type not in (2,7): ok = False
                except:
                    if t in ("title","subtitle","body"): ok = False
                if t == "table":
                    try: ok = shape.HasTable
                    except: ok = False
                if t == "picture" and shape.Type != 13: ok = False
                if t == "chart":
                    try: ok = shape.HasChart
                    except: ok = False
                if t == "textbox":
                    try: ok = shape.Type == 17 and not shape.PlaceholderFormat
                    except: ok = shape.Type == 17
            if ok and "position" in target:
                sw = self.prs.PageSetup.SlideWidth
                sh = self.prs.PageSetup.SlideHeight
                cx = shape.Left + shape.Width/2
                cy = shape.Top + shape.Height/2
                h = "左" if cx < sw*0.33 else ("右" if cx > sw*0.67 else "中")
                v = "上" if cy < sh*0.33 else ("下" if cy > sh*0.67 else "中")
                if target["position"] not in (h+v): ok = False
            if ok and "text_match" in target:
                try:
                    if not shape.HasTextFrame or target["text_match"] not in shape.TextFrame.TextRange.Text:
                        ok = False
                except: ok = False
            if ok and "name" in target:
                try:
                    if target["name"].lower() not in shape.Name.lower():
                        ok = False
                except: ok = False
            if ok and "index" in target:
                if shape_index != target["index"]:
                    ok = False
            if ok: hits.append(shape)
        return hits

    # ---- 文本操作 ----
    def modify_text(self, shape, text):
        old = shape.TextFrame.TextRange.Text
        shape.TextFrame.TextRange.Text = text
        return f"文本: '{old[:20]}' → '{text[:20]}'"

    def modify_partial_text(self, shape, start, length, new_text):
        """修改部分文本 (COM 1-based index)"""
        old = shape.TextFrame.TextRange.Characters(start, length).Text
        shape.TextFrame.TextRange.Characters(start, length).Text = new_text
        return f"部分文本: '{old}' → '{new_text}'"

    def modify_font(self, shape, **kw):
        tr = shape.TextFrame.TextRange
        ch = []
        if "font_size" in kw:       tr.Font.Size = kw["font_size"];       ch.append(f"字号→{kw['font_size']}")
        if "font_size_factor" in kw:
            old = tr.Font.Size
            if old > 0:
                new = round(old * kw["font_size_factor"], 1)
                tr.Font.Size = new; ch.append(f"字号 {old}→{new}")
        if "bold" in kw:          tr.Font.Bold = kw["bold"];          ch.append("加粗" if kw["bold"] else "取消加粗")
        if "italic" in kw:        tr.Font.Italic = kw["italic"];      ch.append("斜体")
        if "underline" in kw:     tr.Font.Underline = kw["underline"]; ch.append("下划线" if kw["underline"] else "取消下划线")
        if "strikethrough" in kw:
            try:
                for pi in range(1, tr.Paragraphs().Count + 1):
                    for ri in range(1, tr.Paragraphs(pi).Runs().Count + 1):
                        tr.Paragraphs(pi).Runs(ri).Font.Strikethrough = -1 if kw["strikethrough"] else 0
            except Exception:
                pass  # Strikethrough not supported on all shapes
            ch.append("删除线" if kw["strikethrough"] else "取消删除线")
        if "color" in kw:         tr.Font.Color.RGB = kw["color"];    ch.append(f"颜色→{hex(kw['color'])}")
        if "font_name" in kw:     tr.Font.Name = kw["font_name"];    ch.append(f"字体→{kw['font_name']}")
        return ", ".join(ch)

    def set_alignment(self, shape, align):
        """设置段落对齐: 左=1, 居中=2, 右=3, 两端=4"""
        align_map = {"左": 1, "left": 1, "居中": 2, "center": 2, "右": 3, "right": 3, "两端": 4, "justify": 4}
        val = align_map.get(align, align) if isinstance(align, str) else align
        for pi in range(1, shape.TextFrame.TextRange.Paragraphs().Count + 1):
            shape.TextFrame.TextRange.Paragraphs(pi).ParagraphFormat.Alignment = val
        return f"对齐方式 → {align}"

    # ---- 形状外观 ----
    def set_fill(self, shape, color_bgr):
        """设置形状填充颜色 (BGR)"""
        shape.Fill.Solid()
        shape.Fill.ForeColor.RGB = color_bgr
        return f"填充颜色 → {hex(color_bgr)}"

    def set_border(self, shape, color_bgr=None, weight=None):
        """设置形状边框"""
        ch = []
        if color_bgr is not None:
            shape.Line.ForeColor.RGB = color_bgr; ch.append(f"边框颜色→{hex(color_bgr)}")
        if weight is not None:
            shape.Line.Weight = weight; ch.append(f"边框粗细→{weight}")
        return ", ".join(ch) if ch else "边框未修改"

    # ---- 形状位置/大小 ----
    def move_shape(self, shape, left=None, top=None):
        ch = []
        if left is not None: shape.Left = left; ch.append(f"Left→{left}")
        if top is not None:  shape.Top = top;   ch.append(f"Top→{top}")
        return f"移动 [{shape.Name}] {', '.join(ch)}"

    def resize_shape(self, shape, width=None, height=None):
        ch = []
        if width is not None:  shape.Width = width;   ch.append(f"Width→{width}")
        if height is not None: shape.Height = height; ch.append(f"Height→{height}")
        return f"缩放 [{shape.Name}] {', '.join(ch)}"

    def delete_shape(self, shape):
        n = shape.Name; shape.Delete(); return f"删除 [{n}]"

    # ---- 文本框/图片 ----
    def add_textbox(self, slide_idx, text, left=100, top=100, width=300, height=50):
        """添加文本框 (msoTextOrientationHorizontal=1)"""
        slide = self.prs.Slides(slide_idx)
        shape = slide.Shapes.AddTextbox(1, left, top, width, height)
        shape.TextFrame.TextRange.Text = text
        return f"第{slide_idx}页添加文本框: '{text[:30]}'"

    def add_picture(self, slide_idx, pic_path, left=100, top=100, width=200, height=150):
        """插入图片"""
        slide = self.prs.Slides(slide_idx)
        abs_path = os.path.abspath(pic_path)
        slide.Shapes.AddPicture(abs_path, False, True, left, top, width, height)
        return f"第{slide_idx}页插入图片: {pic_path}"

    # ---- 幻灯片管理 ----
    def add_slide(self, index=None, layout=1):
        """添加幻灯片: ppLayoutTitle=1, ppLayoutText=2, ppLayoutBlank=12"""
        if index is None: index = self.prs.Slides.Count + 1
        self.prs.Slides.Add(index, layout)
        return f"添加幻灯片: 第{index}页 (layout={layout})"

    def delete_slide(self, slide_idx):
        self.prs.Slides(slide_idx).Delete()
        return f"删除第{slide_idx}页"

    def move_slide(self, slide_idx, new_pos):
        self.prs.Slides(slide_idx).MoveTo(new_pos)
        return f"第{slide_idx}页移动到第{new_pos}页"

    # ---- 表格操作 ----
    def modify_cell(self, slide_idx, target, row, col, text):
        """修改表格单元格 (1-based)"""
        shapes = self.find_shape(slide_idx, target if target else {"type": "table"})
        if not shapes: return f"第{slide_idx}页未找到表格"
        table = shapes[0].Table
        old = table.Cell(row, col).Shape.TextFrame.TextRange.Text
        table.Cell(row, col).Shape.TextFrame.TextRange.Text = text
        return f"表格({row},{col}): '{old[:20]}' → '{text[:20]}'"

    def add_table_row(self, shape):
        shape.Table.Rows.Add()
        return f"表格添加一行 (共{shape.Table.Rows.Count}行)"

    def delete_table_row(self, shape, row):
        shape.Table.Rows(row).Delete()
        return f"表格删除第{row}行"

    def add_table_column(self, shape):
        shape.Table.Columns.Add()
        return f"表格添加一列 (共{shape.Table.Columns.Count}列)"

    def delete_table_column(self, shape, col):
        shape.Table.Columns(col).Delete()
        return f"表格删除第{col}列"

    # ---- COM 独有: 动画 ----
    def add_animation(self, slide_idx, shape, effect="appear"):
        emap = {"appear":1, "fly":2, "fade":10, "zoom":53, "bounce":26}
        eid = emap.get(effect, 1)
        self.prs.Slides(slide_idx).TimeLine.MainSequence.AddEffect(
            Shape=shape, effectId=eid, trigger=1)
        return f"动画 [{shape.Name}] → {effect}"

    def remove_animation(self, slide_idx, anim_index=None):
        """删除动画: 指定索引删单个，不指定删全部"""
        seq = self.prs.Slides(slide_idx).TimeLine.MainSequence
        if anim_index:
            seq(anim_index).Delete()
            return f"第{slide_idx}页删除第{anim_index}个动画"
        else:
            count = seq.Count
            while seq.Count > 0:
                seq(1).Delete()
            return f"第{slide_idx}页清除所有动画 ({count}个)"

    def modify_animation_effect(self, slide_idx, anim_index, new_effect):
        """修改动画效果类型"""
        emap = {"appear":1, "fly":2, "fade":10, "zoom":53, "bounce":26}
        eid = emap.get(new_effect, new_effect) if isinstance(new_effect, str) else new_effect
        self.prs.Slides(slide_idx).TimeLine.MainSequence(anim_index).EffectType = eid
        return f"第{slide_idx}页第{anim_index}个动画效果 → {new_effect}"

    # ---- COM 独有: 切换/导出 ----
    def set_transition(self, slide_idx, trans="fade", dur=1.0):
        tmap = {"fade":3849, "wipe":769, "split":3073, "none":0, "dissolve":1537, "cut":257, "push":3334, "cover":1025, "uncover":1793, "random":513}
        s = self.prs.Slides(slide_idx)
        s.SlideShowTransition.EntryEffect = tmap.get(trans, 2745)
        s.SlideShowTransition.Duration = dur
        return f"第{slide_idx}页切换 → {trans}"

    def export_pdf(self, out):
        self.prs.SaveAs(os.path.abspath(out), 32)  # ppSaveAsPDF=32
        return f"导出 PDF: {out}"

    def export_image(self, slide_idx, out, w=1920, h=1080):
        self.prs.Slides(slide_idx).Export(os.path.abspath(out), "PNG", w, h)
        return f"第{slide_idx}页导出: {out}"

    # ================================================================
    # ================ NEW METHODS BELOW THIS LINE ===================
    # ================================================================

    # ---- Table: add_table ----
    def add_table(self, slide_idx, rows, cols, left=100, top=100, width=400, height=200):
        """Add a table to a slide. Returns the table shape."""
        slide = self.prs.Slides(slide_idx)
        shape = slide.Shapes.AddTable(rows, cols, left, top, width, height)
        return f"Added {rows}x{cols} table on slide {slide_idx}"

    # ---- Duplicate slide ----
    def duplicate_slide(self, slide_idx):
        """Duplicate a slide (inserts copy right after the original)."""
        self.prs.Slides(slide_idx).Duplicate()
        return f"Duplicated slide {slide_idx}"

    # ---- Slide size ----
    def set_slide_size(self, width, height):
        """Set custom slide size in points."""
        self.prs.PageSetup.SlideWidth = width
        self.prs.PageSetup.SlideHeight = height
        return f"Slide size set to {width}x{height}"

    def set_slide_size_preset(self, preset):
        """Set slide size preset. ppSlideSizeOnScreen=0, ppSlideSizeLetterPaper=1,
        ppSlideSizeA4Paper=3, ppSlideSize35MM=4, ppSlideSizeOverhead=5,
        ppSlideSizeBanner=6, ppSlideSizeCustom=7, ppSlideSizeOnScreen16x9=8"""
        preset_map = {"widescreen": 8, "standard": 0, "a4": 3, "letter": 1, "banner": 6}
        val = preset_map.get(preset, preset) if isinstance(preset, str) else preset
        self.prs.PageSetup.SlideSize = val
        return f"Slide size preset set to {preset}"

    # ---- Background ----
    def set_slide_background(self, slide_idx, color_bgr):
        """Set slide background to a solid color (BGR)."""
        slide = self.prs.Slides(slide_idx)
        slide.FollowMasterBackground = False
        slide.Background.Fill.Solid()
        slide.Background.Fill.ForeColor.RGB = color_bgr
        return f"Slide {slide_idx} background set to {hex(color_bgr)}"

    def set_slide_background_image(self, slide_idx, image_path):
        """Set slide background to an image."""
        slide = self.prs.Slides(slide_idx)
        slide.FollowMasterBackground = False
        slide.Background.Fill.UserPicture(os.path.abspath(image_path))
        return f"Slide {slide_idx} background set to {image_path}"

    # ---- Notes ----
    def get_notes(self, slide_idx):
        """Get speaker notes text for a slide."""
        slide = self.prs.Slides(slide_idx)
        try:
            return slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text
        except Exception:
            return ""

    def set_notes(self, slide_idx, text):
        """Set speaker notes text for a slide."""
        slide = self.prs.Slides(slide_idx)
        slide.NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text = text
        return f"Notes set on slide {slide_idx}"

    def append_notes(self, slide_idx, text, separator="\n"):
        """Append text to speaker notes for a slide."""
        current = self.get_notes(slide_idx)
        next_text = f"{current}{separator}{text}" if current else text
        self.set_notes(slide_idx, next_text)
        return f"Notes appended on slide {slide_idx}"

    # ---- Comments ----
    def add_comment(self, slide_idx, text, author="Author", x=10, y=10):
        """Add a comment to a slide."""
        slide = self.prs.Slides(slide_idx)
        slide.Comments.Add(x, y, author, "", text)
        return f"Comment added on slide {slide_idx} by {author}"

    def get_comments(self, slide_idx):
        """Get all comments on a slide."""
        slide = self.prs.Slides(slide_idx)
        result = []
        for i in range(1, slide.Comments.Count + 1):
            c = slide.Comments(i)
            result.append({"author": c.Author, "text": c.Text, "x": c.Left, "y": c.Top})
        return result

    def delete_comment(self, slide_idx, comment_idx):
        """Delete a comment by index (1-based)."""
        slide = self.prs.Slides(slide_idx)
        slide.Comments(comment_idx).Delete()
        return f"Deleted comment {comment_idx} on slide {slide_idx}"

    # ---- Sections ----
    def add_section(self, name, slide_idx):
        """Add a section starting at slide_idx."""
        self.prs.SectionProperties.AddSection(slide_idx, name)
        return f"Section '{name}' added at slide {slide_idx}"

    def delete_section(self, section_idx):
        """Delete a section (1-based). Does not delete slides."""
        self.prs.SectionProperties.Delete(section_idx, False)
        return f"Section {section_idx} deleted"

    def rename_section(self, section_idx, new_name):
        """Rename a section."""
        self.prs.SectionProperties.Rename(section_idx, new_name)
        return f"Section {section_idx} renamed to '{new_name}'"

    def get_sections(self):
        """Get all sections."""
        sp = self.prs.SectionProperties
        result = []
        for i in range(1, sp.Count + 1):
            result.append({"index": i, "name": sp.Name(i),
                           "first_slide": sp.FirstSlide(i),
                           "slide_count": sp.SlidesCount(i)})
        return result

    # ---- Advanced shape operations ----
    def rotate_shape(self, shape, angle):
        """Rotate shape by angle degrees."""
        shape.Rotation = angle
        return f"Rotated [{shape.Name}] to {angle} degrees"

    def flip_shape(self, shape, direction):
        """Flip shape. direction: 'horizontal' (msoFlipHorizontal=0) or 'vertical' (msoFlipVertical=1)."""
        flip_map = {"horizontal": 0, "vertical": 1}
        val = flip_map.get(direction, direction) if isinstance(direction, str) else direction
        shape.Flip(val)
        return f"Flipped [{shape.Name}] {direction}"

    def set_zorder(self, shape, position):
        """Set z-order. position: 'front'=0(msoBringToFront), 'back'=1(msoSendToBack),
        'forward'=2(msoBringForward), 'backward'=3(msoSendBackward)."""
        zmap = {"front": 0, "back": 1, "forward": 2, "backward": 3}
        val = zmap.get(position, position) if isinstance(position, str) else position
        shape.ZOrder(val)
        return f"Z-order [{shape.Name}] -> {position}"

    def group_shapes(self, slide_idx, shape_names):
        """Group shapes by name on a slide. Returns the grouped shape."""
        slide = self.prs.Slides(slide_idx)
        sr = slide.Shapes.Range(shape_names)
        grp = sr.Group()
        return f"Grouped {shape_names} on slide {slide_idx}"

    def ungroup_shapes(self, shape):
        """Ungroup a grouped shape."""
        name = shape.Name
        shape.Ungroup()
        return f"Ungrouped [{name}]"

    def add_connector(self, slide_idx, shape1, shape2, connector_type=1):
        """Add a connector between two shapes.
        connector_type: msoConnectorStraight=1, msoConnectorElbow=2, msoConnectorCurve=3."""
        slide = self.prs.Slides(slide_idx)
        conn = slide.Shapes.AddConnector(connector_type, 0, 0, 100, 100)
        conn.ConnectorFormat.BeginConnect(shape1, 1)
        conn.ConnectorFormat.EndConnect(shape2, 1)
        conn.RerouteConnections()
        return f"Connector added between [{shape1.Name}] and [{shape2.Name}]"

    def add_freeform(self, slide_idx, points):
        """Add a freeform shape from a list of (x, y) points."""
        slide = self.prs.Slides(slide_idx)
        if len(points) < 2:
            return "Need at least 2 points"
        builder = slide.Shapes.BuildFreeform(0, points[0][0], points[0][1])  # msoEditingAuto=0
        for x, y in points[1:]:
            builder.AddNodes(0, 0, x, y)  # msoSegmentLine=0, msoEditingAuto=0
        shape = builder.ConvertToShape()
        return f"Freeform added on slide {slide_idx} with {len(points)} points"

    # ---- Advanced picture operations ----
    def crop_picture(self, shape, left=0, top=0, right=0, bottom=0):
        """Crop a picture shape. Values are proportions (0.0-1.0)."""
        pf = shape.PictureFormat
        pf.CropLeft = left
        pf.CropTop = top
        pf.CropRight = right
        pf.CropBottom = bottom
        return f"Cropped [{shape.Name}] L={left} T={top} R={right} B={bottom}"

    def set_brightness(self, shape, val):
        """Set picture brightness (-1.0 to 1.0)."""
        shape.PictureFormat.Brightness = val
        return f"Brightness [{shape.Name}] -> {val}"

    def set_contrast(self, shape, val):
        """Set picture contrast (-1.0 to 1.0)."""
        shape.PictureFormat.Contrast = val
        return f"Contrast [{shape.Name}] -> {val}"

    def replace_picture(self, shape, new_path):
        """Replace picture by deleting and re-inserting at same position/size."""
        slide_idx = shape.Parent.SlideIndex
        left, top, w, h = shape.Left, shape.Top, shape.Width, shape.Height
        shape.Delete()
        slide = self.prs.Slides(slide_idx)
        slide.Shapes.AddPicture(os.path.abspath(new_path), False, True, left, top, w, h)
        return f"Replaced picture on slide {slide_idx}"

    # ---- Advanced text operations ----
    def add_bullet(self, shape, level=1):
        """Set bullet indent level (0-based) on last paragraph."""
        cnt = shape.TextFrame.TextRange.Paragraphs().Count
        shape.TextFrame.TextRange.Paragraphs(cnt).IndentLevel = level
        return f"Bullet level {level} set on [{shape.Name}]"

    def set_text_autofit(self, shape, mode):
        """Set text autofit. mode: 'none'=0(ppAutoSizeNone), 'fit'=1(ppAutoSizeShapeToFitText),
        'shrink'=2(ppAutoSizeMixed)."""
        mode_map = {"none": 0, "fit": 1, "shrink": 2}
        val = mode_map.get(mode, mode) if isinstance(mode, str) else mode
        shape.TextFrame.AutoSize = val
        return f"Text autofit [{shape.Name}] -> {mode}"

    def add_hyperlink(self, shape, url, text=None):
        """Add a hyperlink to a shape's text."""
        tr = shape.TextFrame.TextRange
        if text:
            tr.Text = text
        tr.ActionSettings(1).Hyperlink.Address = url  # ppMouseClick=1
        return f"Hyperlink added to [{shape.Name}]: {url}"

    def set_word_art(self, shape, style):
        """Set WordArt style on shape text. style is an integer (msoTextEffect constants 0-49)."""
        shape.TextFrame.TextRange.Font.WordArt = True
        try:
            shape.TextEffect.PresetTextEffect = style
        except Exception:
            pass
        return f"WordArt style {style} set on [{shape.Name}]"

    def set_line_spacing(self, shape, spacing):
        """Set line spacing for all paragraphs (in points)."""
        for pi in range(1, shape.TextFrame.TextRange.Paragraphs().Count + 1):
            shape.TextFrame.TextRange.Paragraphs(pi).ParagraphFormat.SpaceWithin = spacing
        return f"Line spacing [{shape.Name}] -> {spacing}"

    def set_paragraph_spacing(self, shape, before=0, after=0):
        """Set paragraph spacing before/after (in points)."""
        for pi in range(1, shape.TextFrame.TextRange.Paragraphs().Count + 1):
            pf = shape.TextFrame.TextRange.Paragraphs(pi).ParagraphFormat
            pf.SpaceBefore = before
            pf.SpaceAfter = after
        return f"Paragraph spacing [{shape.Name}] before={before} after={after}"

    # ---- Audio/Video ----
    def add_audio(self, slide_idx, audio_path, left=100, top=100, width=50, height=50):
        """Insert an audio file on a slide."""
        slide = self.prs.Slides(slide_idx)
        abs_path = os.path.abspath(audio_path)
        shape = slide.Shapes.AddMediaObject2(abs_path, False, True, left, top, width, height)
        return f"Audio added on slide {slide_idx}: {audio_path}"

    def add_video(self, slide_idx, video_path, left=100, top=100, width=400, height=300):
        """Insert a video file on a slide."""
        slide = self.prs.Slides(slide_idx)
        abs_path = os.path.abspath(video_path)
        shape = slide.Shapes.AddMediaObject2(abs_path, False, True, left, top, width, height)
        return f"Video added on slide {slide_idx}: {video_path}"

    def set_media_playback(self, shape, auto_play=False, loop=False, hide_on_stop=False):
        """Set media playback options."""
        anim = shape.AnimationSettings
        anim.PlaySettings.PlayOnEntry = auto_play
        anim.PlaySettings.LoopUntilStopped = loop
        anim.PlaySettings.HideWhileNotPlaying = hide_on_stop
        return f"Media playback [{shape.Name}] auto={auto_play} loop={loop} hide={hide_on_stop}"

    # ---- Chart operations ----
    def add_chart(self, slide_idx, chart_type=4, data=None, left=100, top=100, width=400, height=300):
        """Add a chart. chart_type: xlColumnClustered=51, xlLine=4, xlPie=5, xlBarClustered=57, xl3DColumn=54.
        data: dict with 'categories' (list) and 'series' (list of {name, values})."""
        slide = self.prs.Slides(slide_idx)
        # ppChartType maps differently; use AddChart2 for Office 2013+
        try:
            chart_shape = slide.Shapes.AddChart2(-1, chart_type, left, top, width, height)
        except Exception:
            chart_shape = slide.Shapes.AddChart(chart_type, left, top, width, height)
        if data:
            try:
                chart = chart_shape.Chart
                wb = chart.ChartData.Workbook
                ws = wb.Worksheets(1)
                cats = data.get("categories", [])
                series_list = data.get("series", [])
                for i, cat in enumerate(cats):
                    ws.Cells(i + 2, 1).Value = cat
                for si, s in enumerate(series_list):
                    ws.Cells(1, si + 2).Value = s.get("name", f"Series{si+1}")
                    for vi, v in enumerate(s.get("values", [])):
                        ws.Cells(vi + 2, si + 2).Value = v
                chart.ChartData.Activate()
                wb.Close(True)
            except Exception as ex:
                pass  # chart data setting may fail in some COM configurations
        return f"Chart added on slide {slide_idx} type={chart_type}"

    def modify_chart_data(self, shape, series_idx, values):
        """Modify chart series data. series_idx is 1-based."""
        chart = shape.Chart
        series = chart.SeriesCollection(series_idx)
        series.Values = values
        return f"Chart series {series_idx} data updated"

    def set_chart_title(self, shape, title):
        """Set chart title."""
        chart = shape.Chart
        chart.HasTitle = True
        chart.ChartTitle.Text = title
        return f"Chart title set to '{title}'"

    def set_chart_style(self, shape, style_id):
        """Set chart style (1-48)."""
        shape.Chart.ChartStyle = style_id
        return f"Chart style set to {style_id}"

    # ---- SmartArt ----
    def add_smartart(self, slide_idx, layout_id=None, left=100, top=100, width=400, height=300):
        """Add SmartArt. layout_id should be a SmartArt layout ID string.
        If None, uses the first available layout."""
        slide = self.prs.Slides(slide_idx)
        try:
            smart_art_layouts = self.app.SmartArtLayouts
            if layout_id is None:
                layout = smart_art_layouts(1)
            elif isinstance(layout_id, int):
                layout = smart_art_layouts(layout_id)
            else:
                layout = smart_art_layouts(layout_id)
            shape = slide.Shapes.AddSmartArt(layout, left, top, width, height)
            return f"SmartArt added on slide {slide_idx}"
        except Exception as ex:
            return f"SmartArt failed: {ex}"

    # ---- Master/Layout ----
    def get_slide_masters(self):
        """Get all slide masters."""
        result = []
        for i in range(1, self.prs.SlideMaster.CustomLayouts.Count + 1):
            result.append({"index": i, "name": self.prs.SlideMaster.CustomLayouts(i).Name})
        return result

    def get_slide_layouts(self):
        """Get all slide layouts from the first master."""
        result = []
        try:
            layouts = self.prs.SlideMaster.CustomLayouts
            for i in range(1, layouts.Count + 1):
                result.append({"index": i, "name": layouts(i).Name})
        except Exception:
            # Fallback: try Designs
            for di in range(1, self.prs.Designs.Count + 1):
                d = self.prs.Designs(di)
                for li in range(1, d.SlideMaster.CustomLayouts.Count + 1):
                    result.append({"design": di, "index": li,
                                   "name": d.SlideMaster.CustomLayouts(li).Name})
        return result

    def set_slide_layout(self, slide_idx, layout_name_or_idx):
        """Set layout for a slide by name or index."""
        slide = self.prs.Slides(slide_idx)
        layouts = self.prs.SlideMaster.CustomLayouts
        if isinstance(layout_name_or_idx, int):
            slide.CustomLayout = layouts(layout_name_or_idx)
        else:
            for i in range(1, layouts.Count + 1):
                if layouts(i).Name == layout_name_or_idx:
                    slide.CustomLayout = layouts(i)
                    break
        return f"Slide {slide_idx} layout set to {layout_name_or_idx}"

    def modify_master_element(self, master_idx, shape_name, **kw):
        """Modify a shape on a slide master by name."""
        master = self.prs.Designs(master_idx).SlideMaster
        for shape in master.Shapes:
            if shape.Name == shape_name:
                if "text" in kw and shape.HasTextFrame:
                    shape.TextFrame.TextRange.Text = kw["text"]
                if "font_size" in kw and shape.HasTextFrame:
                    shape.TextFrame.TextRange.Font.Size = kw["font_size"]
                return f"Master element '{shape_name}' modified"
        return f"Shape '{shape_name}' not found on master {master_idx}"

    # ---- Theme ----
    def apply_theme(self, theme_path):
        """Apply a theme (.thmx) file to the presentation."""
        self.prs.ApplyTheme(os.path.abspath(theme_path))
        return f"Theme applied: {theme_path}"

    def get_theme_colors(self):
        """Get theme color scheme info."""
        try:
            tc = self.prs.SlideMaster.Theme.ThemeColorScheme
            result = []
            for i in range(1, tc.Count + 1):
                result.append({"index": i, "rgb": tc(i).RGB})
            return result
        except Exception as ex:
            return f"Could not get theme colors: {ex}"

    def set_theme_colors(self, color_scheme):
        """Set theme colors. color_scheme is a dict of {index: rgb_bgr_value}."""
        tc = self.prs.SlideMaster.Theme.ThemeColorScheme
        for idx, rgb in color_scheme.items():
            tc(int(idx)).RGB = rgb
        return f"Theme colors updated"

    # ---- 3D / Visual effects ----
    def set_shadow(self, shape, preset):
        """Set shadow preset (0-20+). 0 = no shadow."""
        shape.Shadow.Type = 1 if preset > 0 else 0  # msoShadow1
        if preset > 0:
            try:
                shape.Shadow.Style = preset
            except Exception:
                shape.Shadow.Visible = True
        else:
            shape.Shadow.Visible = False
        return f"Shadow preset {preset} set on [{shape.Name}]"

    def set_reflection(self, shape, preset):
        """Set reflection preset."""
        try:
            shape.Reflection.Type = preset
        except Exception:
            return f"Reflection not supported on [{shape.Name}]"
        return f"Reflection preset {preset} set on [{shape.Name}]"

    def set_glow(self, shape, color_bgr, radius=10):
        """Set glow effect."""
        try:
            shape.Glow.Color.RGB = color_bgr
            shape.Glow.Radius = radius
        except Exception:
            return f"Glow not supported on [{shape.Name}]"
        return f"Glow set on [{shape.Name}] color={hex(color_bgr)} radius={radius}"

    def set_3d_rotation(self, shape, x=0, y=0, z=0):
        """Set 3D rotation on shape."""
        try:
            shape.ThreeD.RotationX = x
            shape.ThreeD.RotationY = y
            shape.ThreeD.RotationZ = z
        except Exception:
            return f"3D rotation not supported on [{shape.Name}]"
        return f"3D rotation [{shape.Name}] x={x} y={y} z={z}"

    # ---- Print ----
    def print_presentation(self, printer_name=None, copies=1, print_range=None):
        """Print the presentation."""
        if printer_name:
            self.app.ActivePrinter = printer_name
        if print_range:
            self.prs.PrintOptions.RangeType = 4  # ppPrintSlideRange
            self.prs.PrintOptions.Ranges.Add(print_range[0], print_range[1])
        self.prs.PrintOut(Copies=copies)
        return f"Printed {copies} copies"

    # ---- SlideShow ----
    def start_slideshow(self, from_slide=1, to_slide=None):
        """Start a slideshow."""
        ss = self.prs.SlideShowSettings
        ss.StartingSlide = from_slide
        if to_slide:
            ss.EndingSlide = to_slide
        ss.Run()
        return f"Slideshow started from slide {from_slide}"

    def set_slideshow_settings(self, loop=False, show_narration=True, show_animation=True):
        """Configure slideshow settings."""
        ss = self.prs.SlideShowSettings
        ss.LoopUntilStopped = loop
        ss.ShowWithNarration = show_narration
        ss.ShowWithAnimation = show_animation
        return f"Slideshow settings: loop={loop} narration={show_narration} animation={show_animation}"

    # ---- Merge presentations ----
    def merge_presentations(self, file_paths, output_path=None):
        """Merge slides from multiple presentations into the current one."""
        for fp in file_paths:
            abs_fp = os.path.abspath(fp)
            idx = self.prs.Slides.Count
            self.prs.Slides.InsertFromFile(abs_fp, idx)
        if output_path:
            self._save_presentation(os.path.abspath(output_path))
        return f"Merged {len(file_paths)} presentations"


# ========== 意图解析 ==========
def parse_intent(instruction):
    intents = []
    slide_num = None
    m = re.search(r'第(\d+)[页张]', instruction)
    if m: slide_num = int(m.group(1))

    target = {}
    if any(w in instruction for w in ["副标题","subtitle"]): target["type"] = "subtitle"
    elif any(w in instruction for w in ["标题","title","题目"]): target["type"] = "title"
    elif any(w in instruction for w in ["正文","内容","body"]): target["type"] = "body"
    elif "表格" in instruction: target["type"] = "table"
    elif any(w in instruction for w in ["图片","picture"]): target["type"] = "picture"

    for pos in ["左上","右上","左下","右下","居中"]:
        if pos in instruction: target["position"] = pos

    for qm in re.finditer(r'[""「](.+?)[""」]', instruction):
        before = instruction[:qm.start()]
        if not re.search(r'(?:改|换|替换|变)[成为]?\s*$', before):
            target["text_match"] = qm.group(1); break

    # 按名称定位
    m = re.search(r'(?:name|名称)[：:]([^\s,，]+)', instruction)
    if m: target["name"] = m.group(1)

    # 按索引定位
    m = re.search(r'#(\d+)', instruction)
    if m: target["index"] = int(m.group(1))
    m = re.search(r'第(\d+)个', instruction)
    if m and "index" not in target: target["index"] = int(m.group(1))

    # ---- 幻灯片管理 (先匹配，避免被"删除"等通用规则吃掉) ----

    # 添加幻灯片
    m = re.search(r'(?:添加|新增|插入)\s*(?:一)?[页张]?\s*(?:幻灯片|PPT)?', instruction)
    if m and not any(w in instruction for w in ["文本框","图片","动画","一行","一列"]):
        layout = 12  # ppLayoutBlank
        if "标题" in instruction: layout = 1
        elif "文本" in instruction or "内容" in instruction: layout = 2
        m2 = re.search(r'(?:在)?第(\d+)[页张](?:后|之后)?', instruction)
        idx = int(m2.group(1)) + 1 if m2 else None
        intents.append({"action":"add_slide","slide":slide_num,"target":target,"params":{"index":idx,"layout":layout}})

    # 删除幻灯片
    if re.search(r'删除第\d+[页张]$', instruction) or re.search(r'删除?\s*第\d+[页张]\s*(?:幻灯片)?$', instruction):
        if slide_num:
            intents.append({"action":"delete_slide","slide":slide_num,"target":target,"params":{}})

    # 移动幻灯片
    m = re.search(r'第(\d+)[页张]\s*移[动到]+\s*第(\d+)[页张]', instruction)
    if m:
        intents.append({"action":"move_slide","slide":int(m.group(1)),"target":target,"params":{"new_pos":int(m.group(2))}})

    # ---- 表格操作 ----

    # 修改单元格
    m = re.search(r'表格\s*第?(\d+)\s*行\s*第?(\d+)\s*列\s*(?:改[成为]?|换[成为]?|设为)?\s*[""「]?(.+?)[""」]?\s*$', instruction)
    if m:
        intents.append({"action":"modify_cell","slide":slide_num,"target":{"type":"table"},
                         "params":{"row":int(m.group(1)),"col":int(m.group(2)),"text":m.group(3)}})

    # 表格添加/删除行
    if re.search(r'表格\s*(?:添加|加|新增)\s*(?:一)?行', instruction):
        intents.append({"action":"table_row_add","slide":slide_num,"target":{"type":"table"},"params":{}})
    m = re.search(r'表格\s*删除?\s*第?(\d+)\s*行', instruction)
    if m and "列" not in instruction:
        intents.append({"action":"table_row_delete","slide":slide_num,"target":{"type":"table"},"params":{"row":int(m.group(1))}})

    # 表格添加/删除列
    if re.search(r'表格\s*(?:添加|加|新增)\s*(?:一)?列', instruction):
        intents.append({"action":"table_col_add","slide":slide_num,"target":{"type":"table"},"params":{}})
    m = re.search(r'表格\s*删除?\s*第?(\d+)\s*列', instruction)
    if m and "行" not in instruction:
        intents.append({"action":"table_col_delete","slide":slide_num,"target":{"type":"table"},"params":{"col":int(m.group(1))}})

    # ---- 添加文本框 ----
    m = re.search(r'(?:添加|加个?)\s*文本框\s*(?:内容[是为]?)?\s*[""「]?(.+?)[""」]?\s*$', instruction)
    if m:
        intents.append({"action":"add_textbox","slide":slide_num or 1,"target":target,"params":{"text":m.group(1)}})

    # ---- 插入图片 ----
    m = re.search(r'(?:插入|添加)\s*图片\s*[""「]?(\S+?\.\w{3,4})[""」]?', instruction)
    if m:
        intents.append({"action":"add_picture","slide":slide_num or 1,"target":target,"params":{"pic_path":m.group(1)}})

    # ---- 对齐 ----
    m = re.search(r'(左|右|居中|两端)\s*对齐', instruction)
    if m:
        intents.append({"action":"set_alignment","slide":slide_num,"target":target,"params":{"align":m.group(1)}})

    # ---- 填充 ----
    for cn, bgr in COLOR_MAP.items():
        if re.search(rf'(?:背景|填充)\s*(?:改[成为]?|换[成为]?)?\s*{cn}', instruction):
            intents.append({"action":"set_fill","slide":slide_num,"target":target,"params":{"color_bgr":bgr}})
            break

    # ---- 边框 ----
    border_params = {}
    for cn, bgr in COLOR_MAP.items():
        if re.search(rf'边框\s*(?:改[成为]?|换[成为]?)?\s*{cn}', instruction):
            border_params["color_bgr"] = bgr; break
    m = re.search(r'边框\s*(?:加粗|粗细|宽度)\s*(?:改[成为]?)?\s*(\d+(?:\.\d+)?)', instruction)
    if m: border_params["weight"] = float(m.group(1))
    elif "边框加粗" in instruction: border_params["weight"] = 3.0
    if border_params:
        intents.append({"action":"set_border","slide":slide_num,"target":target,"params":border_params})

    # ---- 移动形状 ----
    m = re.search(r'(?:移动|位置)\s*(?:到|调到)?\s*\(?\s*(\d+)\s*[,，]\s*(\d+)\s*\)?', instruction)
    if m:
        intents.append({"action":"move_shape","slide":slide_num,"target":target,
                         "params":{"left":int(m.group(1)),"top":int(m.group(2))}})
    elif "移动到左上" in instruction:
        intents.append({"action":"move_shape","slide":slide_num,"target":target,"params":{"left":0,"top":0}})

    # ---- 缩放形状 ----
    m = re.search(r'宽度\s*(?:改[成为]?|调[成为]?|设为)?\s*(\d+)', instruction)
    if m:
        intents.append({"action":"resize_shape","slide":slide_num,"target":target,"params":{"width":int(m.group(1))}})
    m = re.search(r'高度\s*(?:改[成为]?|调[成为]?|设为)?\s*(\d+)', instruction)
    if m:
        intents.append({"action":"resize_shape","slide":slide_num,"target":target,"params":{"height":int(m.group(1))}})
    if "放大" in instruction and not any(i["action"] == "resize_shape" for i in intents):
        intents.append({"action":"resize_shape","slide":slide_num,"target":target,"params":{"scale_factor":1.5}})
    elif "缩小" in instruction and not any(i["action"] == "resize_shape" for i in intents):
        intents.append({"action":"resize_shape","slide":slide_num,"target":target,"params":{"scale_factor":0.7}})

    # ---- 删除/清除动画 ----
    if re.search(r'(?:删除|清除|去掉)\s*(?:所有)?\s*动画', instruction):
        intents.append({"action":"remove_animation","slide":slide_num,"target":target,"params":{}})

    # ---- 修改文本 ----
    m = re.search(r'(?:改|换|替换|变)[成为]?\s*[""「](.+?)[""」]', instruction)
    if m and not any(i["action"] in ("modify_cell","add_textbox") for i in intents):
        intents.append({"action":"modify_text","slide":slide_num,"target":target,"params":{"new_text":m.group(1)}})

    # ---- 字号 ----
    m = re.search(r'字号?\s*(?:改[成为]?|调[成为]?|设为)?\s*(\d+)', instruction)
    if m: intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"font_size":int(m.group(1))}})
    elif "大一点" in instruction: intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"font_size_factor":1.3}})
    elif "小一点" in instruction: intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"font_size_factor":0.75}})

    # 加粗
    if any(w in instruction for w in ["加粗","粗体","bold"]):
        intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"bold":True}})

    # 下划线
    if any(w in instruction for w in ["下划线","underline"]):
        intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"underline":True}})

    # 删除线
    if any(w in instruction for w in ["删除线","删除线效果","strikethrough"]):
        intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"strikethrough":True}})

    # 颜色（字体颜色，非填充/边框）
    if not any(i["action"] in ("set_fill","set_border") for i in intents):
        for cn, bgr in COLOR_MAP.items():
            if cn in instruction:
                if re.search(rf'(?:改|换|变|调)[成为]?\s*{cn}', instruction):
                    intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"color":bgr}})
                break

    # 删除元素（排除已匹配的删除幻灯片/动画/表格行列）
    if any(w in instruction for w in ["删除","删掉","去掉"]):
        if not any(i["action"] in ("delete_slide","remove_animation","table_row_delete","table_col_delete") for i in intents):
            intents.append({"action":"delete","slide":slide_num,"target":target,"params":{}})

    # COM 独有: 添加动画
    m = re.search(r'(?:添加|加|设置)\s*动画\s*(\S+)?', instruction)
    if m:
        ef = m.group(1) or "appear"
        cn_map = {"淡入":"fade","飞入":"fly","出现":"appear","缩放":"zoom"}
        intents.append({"action":"animation","slide":slide_num,"target":target,"params":{"effect":cn_map.get(ef,ef)}})

    # COM 独有: 切换效果
    m = re.search(r'切换\s*(?:效果)?\s*(\S+)?', instruction)
    if m:
        tr = m.group(1) or "fade"
        cn_map = {"淡化":"fade","推入":"push","擦除":"wipe"}
        intents.append({"action":"transition","slide":slide_num,"target":target,"params":{"transition":cn_map.get(tr,tr)}})

    # 导出PDF
    if any(w in instruction for w in ["导出pdf","导出PDF","转pdf","转PDF"]):
        intents.append({"action":"export_pdf","slide":slide_num,"target":target,"params":{}})

    # 字体名称
    m = re.search(r'字体\s*(?:改[成为]?|换[成为]?|用)?\s*(\S+)', instruction)
    if m and m.group(1) not in ["大","小","一点"]:
        intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"font_name":m.group(1)}})

    if not intents:
        intents.append({"action":"unknown","raw":instruction,"target":target,"slide":slide_num})
    return intents


# ========== 执行 ==========
def run(ppt, intents, output):
    changes = []
    for intent in intents:
        a = intent["action"]; t = intent.get("target",{}); p = intent.get("params",{}); sn = intent.get("slide")
        if a == "unknown": print(f"⚠️ 无法理解: {intent.get('raw','')}"); continue
        if a == "export_pdf": changes.append(ppt.export_pdf(output.replace(".pptx",".pdf"))); continue
        if a == "transition":
            rng = [sn] if sn else list(range(1, ppt.prs.Slides.Count+1))
            for si in rng: changes.append(ppt.set_transition(si, p.get("transition","fade")))
            continue

        # 幻灯片管理 (不需要查找shape)
        if a == "add_slide":
            changes.append(f"✅ {ppt.add_slide(p.get('index'), p.get('layout',1))}")
            continue
        if a == "delete_slide":
            changes.append(f"✅ {ppt.delete_slide(sn)}")
            continue
        if a == "move_slide":
            changes.append(f"✅ {ppt.move_slide(sn, p['new_pos'])}")
            continue

        # 添加文本框 (不需要查找shape)
        if a == "add_textbox":
            si = sn or 1
            changes.append(f"✅ {ppt.add_textbox(si, p['text'], p.get('left',100), p.get('top',100), p.get('width',300), p.get('height',50))}")
            continue

        # 插入图片 (不需要查找shape)
        if a == "add_picture":
            si = sn or 1
            changes.append(f"✅ {ppt.add_picture(si, p['pic_path'], p.get('left',100), p.get('top',100), p.get('width',200), p.get('height',150))}")
            continue

        # 修改单元格 (用modify_cell自带find_shape)
        if a == "modify_cell":
            si = sn or 1
            changes.append(f"✅ 第{si}页 {ppt.modify_cell(si, t, p['row'], p['col'], p['text'])}")
            continue

        # 删除/清除动画
        if a == "remove_animation":
            rng = [sn] if sn else list(range(1, ppt.prs.Slides.Count+1))
            for si in rng:
                try: changes.append(f"✅ {ppt.remove_animation(si, p.get('anim_index'))}")
                except Exception as ex: print(f"⚠️ 第{si}页动画操作错误: {ex}")
            continue

        # 需要查找shape的操作
        rng = [sn] if sn else list(range(1, ppt.prs.Slides.Count+1))
        # move_shape/resize_shape: 清除 target 中的 position，避免把目标位置误当定位条件
        if a in ("move_shape", "resize_shape"):
            t.pop("position", None)
        for si in rng:
            shapes = ppt.find_shape(si, t)
            if not shapes:
                if sn: print(f"⚠️ 第{si}页: 没找到")
                continue
            if len(shapes) > 1:
                print(f"ℹ️ 第{si}页: {len(shapes)} 个匹配，全部执行")
            for shape in shapes:
                try:
                    if a == "modify_text":
                        changes.append(f"✅ 第{si}页 {ppt.modify_text(shape, p['new_text'])}")
                    elif a == "modify_font":
                        changes.append(f"✅ 第{si}页 [{shape.Name}] {ppt.modify_font(shape, **p)}")
                    elif a == "delete":
                        changes.append(f"✅ 第{si}页 {ppt.delete_shape(shape)}")
                    elif a == "animation":
                        changes.append(f"✅ 第{si}页 {ppt.add_animation(si, shape, p.get('effect','appear'))}")
                    elif a == "set_alignment":
                        changes.append(f"✅ 第{si}页 [{shape.Name}] {ppt.set_alignment(shape, p['align'])}")
                    elif a == "set_fill":
                        changes.append(f"✅ 第{si}页 [{shape.Name}] {ppt.set_fill(shape, p['color_bgr'])}")
                    elif a == "set_border":
                        changes.append(f"✅ 第{si}页 [{shape.Name}] {ppt.set_border(shape, p.get('color_bgr'), p.get('weight'))}")
                    elif a == "move_shape":
                        changes.append(f"✅ 第{si}页 {ppt.move_shape(shape, p.get('left'), p.get('top'))}")
                    elif a == "resize_shape":
                        if "scale_factor" in p:
                            w = shape.Width * p["scale_factor"]
                            h = shape.Height * p["scale_factor"]
                            changes.append(f"✅ 第{si}页 {ppt.resize_shape(shape, w, h)}")
                        else:
                            changes.append(f"✅ 第{si}页 {ppt.resize_shape(shape, p.get('width'), p.get('height'))}")
                    elif a == "table_row_add":
                        changes.append(f"✅ 第{si}页 {ppt.add_table_row(shape)}")
                    elif a == "table_row_delete":
                        changes.append(f"✅ 第{si}页 {ppt.delete_table_row(shape, p['row'])}")
                    elif a == "table_col_add":
                        changes.append(f"✅ 第{si}页 {ppt.add_table_column(shape)}")
                    elif a == "table_col_delete":
                        changes.append(f"✅ 第{si}页 {ppt.delete_table_column(shape, p['col'])}")
                except Exception as ex:
                    print(f"⚠️ 第{si}页 [{shape.Name}] 错误: {ex}")

    if changes:
        ppt.save(output)
        print(f"\n{'='*50}\n📝 {len(changes)} 项修改:")
        for c in changes: print(f"   {c}")
    else:
        print("\n⚠️ 没有执行修改")
    return changes


# ========== 交互模式 ==========
def interactive(ppt, filepath):
    print(f"\n🎮 交互模式 — 输入自然语言指令，输入 q 退出")
    print(f"   COM 独有: 动画/切换效果/导出PDF/导出图片")
    print(f"   新增: 文本框/图片/幻灯片管理/表格操作/对齐/填充/边框/移动/缩放")
    desc = ppt.inspect()
    ppt.print_structure(desc)
    
    output = filepath.replace(".pptx", "_modified.pptx")
    while True:
        try:
            cmd = input("\n📝 指令> ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if cmd.lower() in ("q","quit","exit"): break
        if cmd == "inspect":
            desc = ppt.inspect(); ppt.print_structure(desc); continue
        if cmd.startswith("export-image"):
            m = re.search(r'(\d+)', cmd)
            si = int(m.group(1)) if m else 1
            ppt.export_image(si, f"slide_{si}.png")
            continue
        
        intents = parse_intent(cmd)
        print(f"🔍 {len(intents)} 个意图: {[i['action'] for i in intents]}")
        run(ppt, intents, output)


# ========== Main ==========
def main():
    if len(sys.argv) < 2:
        print("PPTX COM 编辑器 (Windows + Office)")
        print(f"用法: python {sys.argv[0]} <file.pptx> [指令]")
        print(f"      python {sys.argv[0]} <file.pptx> --inspect")
        print(f"      python {sys.argv[0]} <file.pptx> --interactive")
        print(f"      python {sys.argv[0]} <file.pptx> --export-images")
        print()
        print("COM 独有能力:")
        print('  "给第1页标题添加动画淡入"')
        print('  "第2页切换效果淡化"')
        print('  "导出PDF"')
        print()
        print("新增能力:")
        print('  "添加文本框 内容是xxx"')
        print('  "居中对齐" / "左对齐"')
        print('  "背景改成红色" / "填充蓝色"')
        print('  "添加一页" / "删除第3页"')
        print('  "表格第2行第3列改成xxx"')
        print('  "插入图片 xxx.png"')
        print('  "删除动画" / "清除动画"')
        sys.exit(0)

    path = sys.argv[1]
    if not os.path.exists(path):
        print(f"❌ 文件不存在: {path}"); sys.exit(1)

    ppt = PowerPointCOM(visible=False)
    try:
        ppt.open(path)

        if "--inspect" in sys.argv:
            desc = ppt.inspect(); ppt.print_structure(desc)
            jpath = path.replace(".pptx","_structure.json")
            with open(jpath,"w",encoding="utf-8") as f: json.dump(desc,f,ensure_ascii=False,indent=2,default=str)
            print(f"\n📄 结构导出: {jpath}")
        elif "--interactive" in sys.argv:
            interactive(ppt, path)
        elif "--export-images" in sys.argv:
            for si in range(1, ppt.prs.Slides.Count+1):
                ppt.export_image(si, f"slide_{si}.png")
        elif len(sys.argv) > 2 and not sys.argv[2].startswith("--"):
            args = sys.argv[2:]
            output = None
            if "--output" in args:
                oi = args.index("--output")
                if oi + 1 < len(args):
                    output = args[oi + 1]
                    args = args[:oi] + args[oi+2:]
            instruction = " ".join(args)
            print(f"📝 指令: {instruction}")
            intents = parse_intent(instruction)
            print(f"🔍 {len(intents)} 个意图: {[i['action'] for i in intents]}")
            if not output:
                output = path.replace(".pptx", "_modified.pptx")
            run(ppt, intents, output)
        else:
            desc = ppt.inspect(); ppt.print_structure(desc)
    finally:
        ppt.close()

if __name__ == "__main__":
    main()