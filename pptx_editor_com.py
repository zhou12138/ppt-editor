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

    def open(self, path):
        self.filepath = os.path.abspath(path)
        self.prs = self.app.Presentations.Open(self.filepath, False, False, False)
        print(f"📂 已打开: {self.filepath} ({self.prs.Slides.Count}页)")
        return self

    def close(self):
        try:
            if self.prs: self.prs.Close()
            if self.app: self.app.Quit()
        except: pass
        pythoncom.CoUninitialize()

    def save(self, out=None):
        if out:
            self.prs.SaveAs(os.path.abspath(out))
            print(f"💾 已另存为: {out}")
        else:
            self.prs.Save()
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
                     "text": "", "is_placeholder": False}
                
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
                txt = (e.get("text","") or "(无)")[:40].replace("\n","↵")
                print(f"  [{e['id']}] {e['name']}{ph} ({e['position_label']}) → {txt}")
                if e.get("paragraphs"):
                    p = e["paragraphs"][0]
                    print(f"       字体:{p.get('font')} 字号:{p.get('size')} 粗:{p.get('bold')}")
                if e.get("table"):
                    print(f"       表格: {len(e['table'])}×{len(e['table'][0])}")

    # ---- 元素定位 ----
    def find_shape(self, slide_idx, target):
        slide = self.prs.Slides(slide_idx)
        hits = []
        for shape in slide.Shapes:
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
            if ok: hits.append(shape)
        return hits

    # ---- 操作 ----
    def modify_text(self, shape, text):
        old = shape.TextFrame.TextRange.Text
        shape.TextFrame.TextRange.Text = text
        return f"文本: '{old[:20]}' → '{text[:20]}'"

    def modify_font(self, shape, **kw):
        tr = shape.TextFrame.TextRange
        ch = []
        if "font_size" in kw:       tr.Font.Size = kw["font_size"];       ch.append(f"字号→{kw['font_size']}")
        if "font_size_factor" in kw:
            old = tr.Font.Size
            if old > 0:
                new = round(old * kw["font_size_factor"], 1)
                tr.Font.Size = new; ch.append(f"字号 {old}→{new}")
        if "bold" in kw:     tr.Font.Bold = kw["bold"];     ch.append("加粗" if kw["bold"] else "取消加粗")
        if "italic" in kw:   tr.Font.Italic = kw["italic"]; ch.append("斜体")
        if "color" in kw:    tr.Font.Color.RGB = kw["color"]; ch.append(f"颜色→{hex(kw['color'])}")
        if "font_name" in kw: tr.Font.Name = kw["font_name"]; ch.append(f"字体→{kw['font_name']}")
        return ", ".join(ch)

    def delete_shape(self, shape):
        n = shape.Name; shape.Delete(); return f"删除 [{n}]"

    # ---- COM 独有 ----
    def add_animation(self, slide_idx, shape, effect="appear"):
        emap = {"appear":1, "fly":2, "fade":10, "zoom":53, "bounce":26}
        eid = emap.get(effect, 1)
        self.prs.Slides(slide_idx).TimeLine.MainSequence.AddEffect(
            Shape=shape, effectId=eid, trigger=1)
        return f"动画 [{shape.Name}] → {effect}"

    def set_transition(self, slide_idx, trans="fade", dur=1.0):
        tmap = {"fade":3849, "push":3336, "wipe":769, "split":3073, "none":0, "dissolve":1537, "cut":257}
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

    # 修改文本
    m = re.search(r'(?:改|换|替换|变)[成为]?\s*[""「](.+?)[""」]', instruction)
    if m: intents.append({"action":"modify_text","slide":slide_num,"target":target,"params":{"new_text":m.group(1)}})

    # 字号
    m = re.search(r'字号?\s*(?:改[成为]?|调[成为]?|设为)?\s*(\d+)', instruction)
    if m: intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"font_size":int(m.group(1))}})
    elif "大一点" in instruction: intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"font_size_factor":1.3}})
    elif "小一点" in instruction: intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"font_size_factor":0.75}})

    if any(w in instruction for w in ["加粗","粗体","bold"]):
        intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"bold":True}})

    for cn, bgr in COLOR_MAP.items():
        if cn in instruction:
            if re.search(rf'(?:改|换|变|调)[成为]?\s*{cn}', instruction):
                intents.append({"action":"modify_font","slide":slide_num,"target":target,"params":{"color":bgr}})
            break

    if any(w in instruction for w in ["删除","删掉","去掉"]):
        intents.append({"action":"delete","slide":slide_num,"target":target,"params":{}})

    # COM 独有
    m = re.search(r'(?:添加|加|设置)\s*动画\s*(\S+)?', instruction)
    if m:
        ef = m.group(1) or "appear"
        cn_map = {"淡入":"fade","飞入":"fly","出现":"appear","缩放":"zoom"}
        intents.append({"action":"animation","slide":slide_num,"target":target,"params":{"effect":cn_map.get(ef,ef)}})

    m = re.search(r'切换\s*(?:效果)?\s*(\S+)?', instruction)
    if m:
        tr = m.group(1) or "fade"
        cn_map = {"淡化":"fade","推入":"push","擦除":"wipe"}
        intents.append({"action":"transition","slide":slide_num,"target":target,"params":{"transition":cn_map.get(tr,tr)}})

    if any(w in instruction for w in ["导出pdf","导出PDF","转pdf","转PDF"]):
        intents.append({"action":"export_pdf","slide":slide_num,"target":target,"params":{}})

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

        rng = [sn] if sn else list(range(1, ppt.prs.Slides.Count+1))
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
            instruction = " ".join(sys.argv[2:])
            print(f"📝 指令: {instruction}")
            intents = parse_intent(instruction)
            print(f"🔍 {len(intents)} 个意图: {[i['action'] for i in intents]}")
            output = path.replace(".pptx", "_modified.pptx")
            run(ppt, intents, output)
        else:
            desc = ppt.inspect(); ppt.print_structure(desc)
    finally:
        ppt.close()

if __name__ == "__main__":
    main()
