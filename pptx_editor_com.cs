// PPTX 自然语言编辑器 — C# COM 版本 (Windows + Office)
// 功能对标 pptx_editor_com.py，通过 Microsoft.Office.Interop.PowerPoint 操作
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PptxEditorCom
{
    static class ColorMap
    {
        public static readonly Dictionary<string, int> Map = new Dictionary<string, int>
        {
            {"红",0x0000FF},{"红色",0x0000FF},{"蓝",0xFF0000},{"蓝色",0xFF0000},
            {"绿",0x00AA00},{"绿色",0x00AA00},{"黄",0x00D7FF},{"黄色",0x00D7FF},
            {"黑",0x000000},{"黑色",0x000000},{"白",0xFFFFFF},{"白色",0xFFFFFF},
            {"灰",0x888888},{"灰色",0x888888},{"橙",0x008CFF},{"橙色",0x008CFF},
            {"紫",0x800080},{"紫色",0x800080},{"粉",0xB469FF},{"粉色",0xB469FF},
            {"red",0x0000FF},{"blue",0xFF0000},{"green",0x00AA00},{"black",0x000000},{"white",0xFFFFFF},
        };
    }

    static class PhNames
    {
        public static readonly Dictionary<int, string> Map = new Dictionary<int, string>
        {
            {1,"TITLE"},{2,"BODY"},{3,"CENTER_TITLE"},{4,"SUBTITLE"},{7,"OBJECT"},
            {8,"CHART"},{9,"TABLE"},{12,"MEDIA"},{13,"SLIDE_NUMBER"},{15,"FOOTER"},
        };
        public static string Get(int t) => Map.ContainsKey(t) ? Map[t] : $"({t})";
    }

    class Intent
    {
        public string Action;
        public int? Slide;
        public Dictionary<string, string> Target = new Dictionary<string, string>();
        public Dictionary<string, object> Params = new Dictionary<string, object>();
        public string Raw;
    }

    class PowerPointCOM : IDisposable
    {
        PowerPoint.Application app;
        PowerPoint.Presentation prs;
        string filepath;

        public PowerPointCOM(bool visible = false)
        {
            app = new PowerPoint.Application();
            if (visible) app.Visible = Office.MsoTriState.msoTrue;
        }

        public void Open(string path)
        {
            filepath = Path.GetFullPath(path);
            prs = app.Presentations.Open(filepath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse, Office.MsoTriState.msoFalse);
            Console.WriteLine($"📂 已打开: {filepath} ({prs.Slides.Count}页)");
        }

        public void Close()
        {
            try { if (prs != null) { prs.Close(); RCO(prs); } } catch { }
            try { if (app != null) { app.Quit(); RCO(app); } } catch { }
        }

        public void Dispose() => Close();

        public void Save(string output = null)
        {
            if (output != null) { prs.SaveAs(Path.GetFullPath(output)); Console.WriteLine($"💾 已另存为: {output}"); }
            else { prs.Save(); Console.WriteLine("💾 已保存"); }
        }

        static void RCO(object o) { if (o != null) Marshal.ReleaseComObject(o); }
        float SlideWidth => prs.PageSetup.SlideWidth;
        float SlideHeight => prs.PageSetup.SlideHeight;
        public int SlideCount => prs.Slides.Count;

        string PosLabel(float left, float top, float w, float h)
        {
            float cx = left + w / 2, cy = top + h / 2;
            string hh = cx < SlideWidth * 0.33f ? "左" : (cx > SlideWidth * 0.67f ? "右" : "中");
            string vv = cy < SlideHeight * 0.33f ? "上" : (cy > SlideHeight * 0.67f ? "下" : "中");
            return hh + vv;
        }

        public List<Dictionary<string, object>> Inspect()
        {
            var slides = new List<Dictionary<string, object>>();
            for (int si = 1; si <= prs.Slides.Count; si++)
            {
                var slide = prs.Slides[si];
                var sd = new Dictionary<string, object> { ["index"] = si, ["layout"] = "", ["elements"] = new List<Dictionary<string, object>>() };
                try { sd["layout"] = slide.CustomLayout.Name; } catch { sd["layout"] = slide.Layout.ToString(); }
                var elems = (List<Dictionary<string, object>>)sd["elements"];
                foreach (PowerPoint.Shape shape in slide.Shapes)
                {
                    var e = new Dictionary<string, object>
                    {
                        ["id"] = shape.Id, ["name"] = shape.Name, ["type"] = (int)shape.Type,
                        ["left"] = Math.Round(shape.Left, 1), ["top"] = Math.Round(shape.Top, 1),
                        ["width"] = Math.Round(shape.Width, 1), ["height"] = Math.Round(shape.Height, 1),
                        ["text"] = "", ["is_placeholder"] = false,
                        ["position_label"] = PosLabel(shape.Left, shape.Top, shape.Width, shape.Height),
                    };
                    try { var pf = shape.PlaceholderFormat; if (pf != null) { e["is_placeholder"] = true; e["ph_type"] = (int)pf.Type; e["ph_type_name"] = PhNames.Get((int)pf.Type); } } catch { }
                    try
                    {
                        if (shape.HasTextFrame == Office.MsoTriState.msoTrue)
                        {
                            e["text"] = shape.TextFrame.TextRange.Text;
                            var paras = new List<Dictionary<string, object>>();
                            for (int pi = 1; pi <= shape.TextFrame.TextRange.Paragraphs().Count; pi++)
                            {
                                var p = shape.TextFrame.TextRange.Paragraphs(pi);
                                var pd = new Dictionary<string, object> { ["text"] = p.Text, ["font"] = p.Font.Name, ["size"] = p.Font.Size, ["bold"] = p.Font.Bold == Office.MsoTriState.msoTrue };
                                try { pd["color"] = p.Font.Color.RGB; } catch { }
                                paras.Add(pd);
                            }
                            e["paragraphs"] = paras;
                        }
                    } catch { }
                    try
                    {
                        if (shape.HasTable == Office.MsoTriState.msoTrue)
                        {
                            var t = shape.Table; var table = new List<List<string>>();
                            for (int r = 1; r <= t.Rows.Count; r++) { var row = new List<string>(); for (int c = 1; c <= t.Columns.Count; c++) row.Add(t.Cell(r, c).Shape.TextFrame.TextRange.Text); table.Add(row); }
                            e["table"] = table;
                        }
                    } catch { }
                    elems.Add(e);
                }
                slides.Add(sd);
            }
            return slides;
        }

        public void PrintStructure(List<Dictionary<string, object>> desc)
        {
            foreach (var s in desc)
            {
                Console.WriteLine($"\n==================================================\n📄 第 {s["index"]} 页 ({s["layout"]})\n==================================================");
                foreach (var e in (List<Dictionary<string, object>>)s["elements"])
                {
                    string ph = (bool)e["is_placeholder"] && e.ContainsKey("ph_type_name") ? $" [{e["ph_type_name"]}]" : "";
                    string txt = (string)e["text"]; if (string.IsNullOrEmpty(txt)) txt = "(无)";
                    if (txt.Length > 40) txt = txt.Substring(0, 40); txt = txt.Replace("\n", "↵");
                    Console.WriteLine($"  [{e["id"]}] {e["name"]}{ph} ({e["position_label"]}) → {txt}");
                    if (e.ContainsKey("paragraphs")) { var ps = (List<Dictionary<string, object>>)e["paragraphs"]; if (ps.Count > 0) { var p0 = ps[0]; Console.WriteLine($"       字体:{V(p0,"font")} 字号:{V(p0,"size")} 粗:{V(p0,"bold")}"); } }
                    if (e.ContainsKey("table")) { var tbl = (List<List<string>>)e["table"]; Console.WriteLine($"       表格: {tbl.Count}×{(tbl.Count>0?tbl[0].Count:0)}"); }
                }
            }
        }
        static string V(Dictionary<string, object> d, string k) => d.ContainsKey(k) ? d[k]?.ToString() ?? "" : "";

        public List<PowerPoint.Shape> FindShape(int slideIdx, Dictionary<string, string> target)
        {
            var slide = prs.Slides[slideIdx]; var hits = new List<PowerPoint.Shape>();
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                bool ok = true;
                if (target.ContainsKey("type"))
                {
                    string t = target["type"];
                    try { var pf = shape.PlaceholderFormat; int pt = (int)pf.Type;
                        if (t == "title" && pt != 1 && pt != 3) ok = false;
                        else if (t == "subtitle" && pt != 4) ok = false;
                        else if (t == "body" && pt != 2 && pt != 7) ok = false;
                    } catch { if (t == "title" || t == "subtitle" || t == "body") ok = false; }
                    if (t == "table") { try { ok = shape.HasTable == Office.MsoTriState.msoTrue; } catch { ok = false; } }
                    if (t == "picture" && (int)shape.Type != 13) ok = false;
                }
                if (ok && target.ContainsKey("position")) { string pos = PosLabel(shape.Left, shape.Top, shape.Width, shape.Height); if (!pos.Contains(target["position"])) ok = false; }
                if (ok && target.ContainsKey("text_match")) { try { if (shape.HasTextFrame != Office.MsoTriState.msoTrue || !shape.TextFrame.TextRange.Text.Contains(target["text_match"])) ok = false; } catch { ok = false; } }
                if (ok) hits.Add(shape);
            }
            return hits;
        }

        public string ModifyText(PowerPoint.Shape shape, string text)
        {
            string old = shape.TextFrame.TextRange.Text;
            string os = old.Length > 20 ? old.Substring(0, 20) : old;
            string ns = text.Length > 20 ? text.Substring(0, 20) : text;
            shape.TextFrame.TextRange.Text = text;
            return $"文本: '{os}' → '{ns}'";
        }

        public string ModifyPartialText(PowerPoint.Shape shape, int start, int length, string newText)
        {
            string old = shape.TextFrame.TextRange.Characters(start, length).Text;
            shape.TextFrame.TextRange.Characters(start, length).Text = newText;
            return $"部分文本: '{old}' → '{newText}'";
        }

        public string AddTextbox(int slideIdx, string text, float left = 100, float top = 100, float width = 300, float height = 50)
        {
            var slide = prs.Slides[slideIdx];
            var shape = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height);
            shape.TextFrame.TextRange.Text = text;
            string s = text.Length > 30 ? text.Substring(0, 30) : text;
            return $"第{slideIdx}页添加文本框: '{s}'";
        }

        public string ModifyFont(PowerPoint.Shape shape, Dictionary<string, object> kw)
        {
            var tr = shape.TextFrame.TextRange; var ch = new List<string>();
            if (kw.ContainsKey("font_size")) { tr.Font.Size = Convert.ToSingle(kw["font_size"]); ch.Add($"字号→{kw["font_size"]}"); }
            if (kw.ContainsKey("font_size_factor")) { float old = tr.Font.Size; if (old > 0) { float n = (float)Math.Round(old * Convert.ToSingle(kw["font_size_factor"]), 1); tr.Font.Size = n; ch.Add($"字号 {old}→{n}"); } }
            if (kw.ContainsKey("bold")) { bool v = Convert.ToBoolean(kw["bold"]); tr.Font.Bold = v ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; ch.Add(v ? "加粗" : "取消加粗"); }
            if (kw.ContainsKey("italic")) { bool v = Convert.ToBoolean(kw["italic"]); tr.Font.Italic = v ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; ch.Add("斜体"); }
            if (kw.ContainsKey("underline")) { bool v = Convert.ToBoolean(kw["underline"]); tr.Font.Underline = v ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; ch.Add(v ? "下划线" : "取消下划线"); }
            if (kw.ContainsKey("strikethrough")) { bool v = Convert.ToBoolean(kw["strikethrough"]); tr.Font.Strikethrough = v ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; ch.Add(v ? "删除线" : "取消删除线"); }
            if (kw.ContainsKey("color")) { int c = Convert.ToInt32(kw["color"]); tr.Font.Color.RGB = c; ch.Add($"颜色→0x{c:X}"); }
            if (kw.ContainsKey("font_name")) { tr.Font.Name = (string)kw["font_name"]; ch.Add($"字体→{kw["font_name"]}"); }
            return string.Join(", ", ch);
        }

        public string SetAlignment(PowerPoint.Shape shape, string align)
        {
            var map = new Dictionary<string, int> { {"左",1},{"left",1},{"居中",2},{"center",2},{"右",3},{"right",3},{"两端",4},{"justify",4} };
            int val = map.ContainsKey(align) ? map[align] : 2;
            for (int pi = 1; pi <= shape.TextFrame.TextRange.Paragraphs().Count; pi++)
                shape.TextFrame.TextRange.Paragraphs(pi).ParagraphFormat.Alignment = (PowerPoint.PpParagraphAlignment)val;
            return $"对齐方式 → {align}";
        }

        public string SetFill(PowerPoint.Shape shape, int colorBgr) { shape.Fill.Solid(); shape.Fill.ForeColor.RGB = colorBgr; return $"填充颜色 → 0x{colorBgr:X}"; }

        public string SetBorder(PowerPoint.Shape shape, int? colorBgr = null, float? weight = null)
        {
            var ch = new List<string>();
            if (colorBgr.HasValue) { shape.Line.ForeColor.RGB = colorBgr.Value; ch.Add($"边框颜色→0x{colorBgr.Value:X}"); }
            if (weight.HasValue) { shape.Line.Weight = weight.Value; ch.Add($"边框粗细→{weight.Value}"); }
            return ch.Count > 0 ? string.Join(", ", ch) : "边框未修改";
        }

        public string MoveShape(PowerPoint.Shape shape, float? left = null, float? top = null)
        {
            var ch = new List<string>();
            if (left.HasValue) { shape.Left = left.Value; ch.Add($"Left→{left.Value}"); }
            if (top.HasValue) { shape.Top = top.Value; ch.Add($"Top→{top.Value}"); }
            return $"移动 [{shape.Name}] {string.Join(\", \", ch)}";
        }

        public string ResizeShape(PowerPoint.Shape shape, float? width = null, float? height = null)
        {
            var ch = new List<string>();
            if (width.HasValue) { shape.Width = width.Value; ch.Add($"Width→{width.Value}"); }
            if (height.HasValue) { shape.Height = height.Value; ch.Add($"Height→{height.Value}"); }
            return $"缩放 [{shape.Name}] {string.Join(\", \", ch)}";
        }

        public string DeleteShape(PowerPoint.Shape shape) { string n = shape.Name; shape.Delete(); return $"删除 [{n}]"; }

        public string AddPicture(int slideIdx, string picPath, float left = 100, float top = 100, float width = 200, float height = 150)
        {
            var slide = prs.Slides[slideIdx]; string absPath = Path.GetFullPath(picPath);
            slide.Shapes.AddPicture(absPath, Office.MsoTriState.msoFalse, Office.MsoTriState.msoTrue, left, top, width, height);
            return $"第{slideIdx}页插入图片: {picPath}";
        }

        public string AddSlide(int? index = null, int layout = 1)
        {
            int idx = index ?? prs.Slides.Count + 1;
            prs.Slides.Add(idx, (PowerPoint.PpSlideLayout)layout);
            return $"添加幻灯片: 第{idx}页 (layout={layout})";
        }

        public string DeleteSlide(int slideIdx) { prs.Slides[slideIdx].Delete(); return $"删除第{slideIdx}页"; }
        public string MoveSlide(int slideIdx, int newPos) { prs.Slides[slideIdx].MoveTo(newPos); return $"第{slideIdx}页移动到第{newPos}页"; }

        public string ModifyCell(int slideIdx, Dictionary<string, string> target, int row, int col, string text)
        {
            var tgt = (target != null && target.Count > 0) ? target : new Dictionary<string, string> { { "type", "table" } };
            var shapes = FindShape(slideIdx, tgt);
            if (shapes.Count == 0) return $"第{slideIdx}页未找到表格";
            var table = shapes[0].Table;
            string old = table.Cell(row, col).Shape.TextFrame.TextRange.Text;
            string os = old.Length > 20 ? old.Substring(0, 20) : old;
            string ns = text.Length > 20 ? text.Substring(0, 20) : text;
            table.Cell(row, col).Shape.TextFrame.TextRange.Text = text;
            return $"表格({row},{col}): '{os}' → '{ns}'";
        }

        public string AddTableRow(PowerPoint.Shape shape) { shape.Table.Rows.Add(); return $"表格添加一行 (共{shape.Table.Rows.Count}行)"; }
        public string DeleteTableRow(PowerPoint.Shape shape, int row) { shape.Table.Rows[row].Delete(); return $"表格删除第{row}行"; }
        public string AddTableColumn(PowerPoint.Shape shape) { shape.Table.Columns.Add(); return $"表格添加一列 (共{shape.Table.Columns.Count}列)"; }
        public string DeleteTableColumn(PowerPoint.Shape shape, int col) { shape.Table.Columns[col].Delete(); return $"表格删除第{col}列"; }

        public string AddAnimation(int slideIdx, PowerPoint.Shape shape, string effect = "appear")
        {
            var emap = new Dictionary<string, int> { {"appear",1},{"fly",2},{"fade",10},{"zoom",53},{"bounce",26} };
            int eid = emap.ContainsKey(effect) ? emap[effect] : 1;
            prs.Slides[slideIdx].TimeLine.MainSequence.AddEffect(shape, (PowerPoint.MsoAnimEffect)eid, 0, PowerPoint.MsoAnimTriggerType.msoAnimTriggerOnPageClick);
            return $"动画 [{shape.Name}] → {effect}";
        }

        public string RemoveAnimation(int slideIdx, int? animIndex = null)
        {
            var seq = prs.Slides[slideIdx].TimeLine.MainSequence;
            if (animIndex.HasValue) { seq[animIndex.Value].Delete(); return $"第{slideIdx}页删除第{animIndex.Value}个动画"; }
            int count = seq.Count; while (seq.Count > 0) seq[1].Delete();
            return $"第{slideIdx}页清除所有动画 ({count}个)";
        }

        public string ModifyAnimationEffect(int slideIdx, int animIndex, string newEffect)
        {
            var emap = new Dictionary<string, int> { {"appear",1},{"fly",2},{"fade",10},{"zoom",53},{"bounce",26} };
            int eid = emap.ContainsKey(newEffect) ? emap[newEffect] : 1;
            prs.Slides[slideIdx].TimeLine.MainSequence[animIndex].EffectType = (PowerPoint.MsoAnimEffect)eid;
            return $"第{slideIdx}页第{animIndex}个动画效果 → {newEffect}";
        }

        public string SetTransition(int slideIdx, string trans = "fade", float dur = 1.0f)
        {
            var tmap = new Dictionary<string, int> { {"fade",3849},{"push",3336},{"wipe",769},{"split",3073},{"none",0},{"dissolve",1537},{"cut",257} };
            var s = prs.Slides[slideIdx];
            s.SlideShowTransition.EntryEffect = (PowerPoint.PpEntryEffect)(tmap.ContainsKey(trans) ? tmap[trans] : 2745);
            s.SlideShowTransition.Duration = dur;
            return $"第{slideIdx}页切换 → {trans}";
        }

        public string ExportPdf(string output) { prs.SaveAs(Path.GetFullPath(output), PowerPoint.PpSaveAsFileType.ppSaveAsPDF); return $"导出 PDF: {output}"; }
        public string ExportImage(int slideIdx, string output, int w = 1920, int h = 1080) { prs.Slides[slideIdx].Export(Path.GetFullPath(output), "PNG", w, h); return $"第{slideIdx}页导出: {output}"; }
    }

    // ========== IntentParser ==========
    static class IntentParser
    {
        static bool CA(string s, params string[] w) => w.Any(x => s.Contains(x));
        static Dictionary<string, string> Cl(Dictionary<string, string> d) => new Dictionary<string, string>(d);

        public static List<Intent> ParseIntent(string instruction)
        {
            var intents = new List<Intent>();
            int? sn = null;
            var ms = Regex.Match(instruction, @"\u7b2c(\d+)[\u9875\u5f20]");
            if (ms.Success) sn = int.Parse(ms.Groups[1].Value);

            var target = new Dictionary<string, string>();
            if (CA(instruction, "\u526f\u6807\u9898", "subtitle")) target["type"] = "subtitle";
            else if (CA(instruction, "\u6807\u9898", "title", "\u9898\u76ee")) target["type"] = "title";
            else if (CA(instruction, "\u6b63\u6587", "\u5185\u5bb9", "body")) target["type"] = "body";
            else if (instruction.Contains("\u8868\u683c")) target["type"] = "table";
            else if (CA(instruction, "\u56fe\u7247", "picture")) target["type"] = "picture";

            foreach (var pos in new[] { "\u5de6\u4e0a", "\u53f3\u4e0a", "\u5de6\u4e0b", "\u53f3\u4e0b", "\u5c45\u4e2d" })
                if (instruction.Contains(pos)) { target["position"] = pos; break; }

            foreach (Match qm in Regex.Matches(instruction, "[\u201c\u201d\u300c](.+?)[\u201c\u201d\u300d]"))
            {
                string before = instruction.Substring(0, qm.Index);
                if (!Regex.IsMatch(before, @"(?:\u6539|\u6362|\u66ff\u6362|\u53d8)[\u6210\u4e3a]?\s*$"))
                { target["text_match"] = qm.Groups[1].Value; break; }
            }

            if (Regex.IsMatch(instruction, @"(?:\u6dfb\u52a0|\u65b0\u589e|\u63d2\u5165)\s*(?:\u4e00)?[\u9875\u5f20]?\s*(?:\u5e7b\u706f\u7247|PPT)?") &&
                !CA(instruction, "\u6587\u672c\u6846", "\u56fe\u7247", "\u52a8\u753b", "\u4e00\u884c", "\u4e00\u5217"))
            {
                int layout = 12;
                if (instruction.Contains("\u6807\u9898")) layout = 1;
                else if (CA(instruction, "\u6587\u672c", "\u5185\u5bb9")) layout = 2;
                var m2 = Regex.Match(instruction, @"(?:\u5728)?\u7b2c(\d+)[\u9875\u5f20](?:\u540e|\u4e4b\u540e)?");
                object idx = m2.Success ? (object)(int.Parse(m2.Groups[1].Value) + 1) : null;
                intents.Add(new Intent { Action = "add_slide", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["index"] = idx, ["layout"] = layout } });
            }

            if ((Regex.IsMatch(instruction, @"\u5220\u9664\u7b2c\d+[\u9875\u5f20]$") ||
                 Regex.IsMatch(instruction, @"\u5220\u9664?\s*\u7b2c\d+[\u9875\u5f20]\s*(?:\u5e7b\u706f\u7247)?$")) && sn.HasValue)
                intents.Add(new Intent { Action = "delete_slide", Slide = sn, Target = Cl(target) });

            var mMove = Regex.Match(instruction, @"\u7b2c(\d+)[\u9875\u5f20]\s*\u79fb[\u52a8\u5230]+\s*\u7b2c(\d+)[\u9875\u5f20]");
            if (mMove.Success)
                intents.Add(new Intent { Action = "move_slide", Slide = int.Parse(mMove.Groups[1].Value), Target = Cl(target),
                    Params = new Dictionary<string, object> { ["new_pos"] = int.Parse(mMove.Groups[2].Value) } });

            var mCell = Regex.Match(instruction, @"\u8868\u683c\s*\u7b2c?(\d+)\s*\u884c\s*\u7b2c?(\d+)\s*\u5217\s*(?:\u6539[\u6210\u4e3a]?|\u6362[\u6210\u4e3a]?|\u8bbe\u4e3a)?\s*[\u201c\u201d\u300c]?(.+?)[\u201c\u201d\u300d]?\s*$");
            if (mCell.Success)
                intents.Add(new Intent { Action = "modify_cell", Slide = sn, Target = new Dictionary<string, string> { { "type", "table" } },
                    Params = new Dictionary<string, object> { ["row"] = int.Parse(mCell.Groups[1].Value), ["col"] = int.Parse(mCell.Groups[2].Value), ["text"] = mCell.Groups[3].Value } });

            if (Regex.IsMatch(instruction, @"\u8868\u683c\s*(?:\u6dfb\u52a0|\u52a0|\u65b0\u589e)\s*(?:\u4e00)?\u884c"))
                intents.Add(new Intent { Action = "table_row_add", Slide = sn, Target = new Dictionary<string, string> { { "type", "table" } } });
            var mTRD = Regex.Match(instruction, @"\u8868\u683c\s*\u5220\u9664?\s*\u7b2c?(\d+)\s*\u884c");
            if (mTRD.Success && !instruction.Contains("\u5217"))
                intents.Add(new Intent { Action = "table_row_delete", Slide = sn, Target = new Dictionary<string, string> { { "type", "table" } },
                    Params = new Dictionary<string, object> { ["row"] = int.Parse(mTRD.Groups[1].Value) } });

            if (Regex.IsMatch(instruction, @"\u8868\u683c\s*(?:\u6dfb\u52a0|\u52a0|\u65b0\u589e)\s*(?:\u4e00)?\u5217"))
                intents.Add(new Intent { Action = "table_col_add", Slide = sn, Target = new Dictionary<string, string> { { "type", "table" } } });
            var mTCD = Regex.Match(instruction, @"\u8868\u683c\s*\u5220\u9664?\s*\u7b2c?(\d+)\s*\u5217");
            if (mTCD.Success && !instruction.Contains("\u884c"))
                intents.Add(new Intent { Action = "table_col_delete", Slide = sn, Target = new Dictionary<string, string> { { "type", "table" } },
                    Params = new Dictionary<string, object> { ["col"] = int.Parse(mTCD.Groups[1].Value) } });

            var mTB = Regex.Match(instruction, @"(?:\u6dfb\u52a0|\u52a0\u4e2a?)\s*\u6587\u672c\u6846\s*(?:\u5185\u5bb9[\u662f\u4e3a]?)?\s*[\u201c\u201d\u300c]?(.+?)[\u201c\u201d\u300d]?\s*$");
            if (mTB.Success)
                intents.Add(new Intent { Action = "add_textbox", Slide = sn ?? 1, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["text"] = mTB.Groups[1].Value } });

            var mPic = Regex.Match(instruction, @"(?:\u63d2\u5165|\u6dfb\u52a0)\s*\u56fe\u7247\s*[\u201c\u201d\u300c]?(\S+?\.\w{3,4})[\u201c\u201d\u300d]?");
            if (mPic.Success)
                intents.Add(new Intent { Action = "add_picture", Slide = sn ?? 1, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["pic_path"] = mPic.Groups[1].Value } });

            var mAl = Regex.Match(instruction, @"(\u5de6|\u53f3|\u5c45\u4e2d|\u4e24\u7aef)\s*\u5bf9\u9f50");
            if (mAl.Success)
                intents.Add(new Intent { Action = "set_alignment", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["align"] = mAl.Groups[1].Value } });

            foreach (var kv in ColorMap.Map)
            {
                if (Regex.IsMatch(instruction, $@"(?:\u80cc\u666f|\u586b\u5145)\s*(?:\u6539[\u6210\u4e3a]?|\u6362[\u6210\u4e3a]?)?\s*{Regex.Escape(kv.Key)}"))
                { intents.Add(new Intent { Action = "set_fill", Slide = sn, Target = Cl(target), Params = new Dictionary<string, object> { ["color_bgr"] = kv.Value } }); break; }
            }

            var borderParams = new Dictionary<string, object>();
            foreach (var kv in ColorMap.Map)
            {
                if (Regex.IsMatch(instruction, $@"\u8fb9\u6846\s*(?:\u6539[\u6210\u4e3a]?|\u6362[\u6210\u4e3a]?)?\s*{Regex.Escape(kv.Key)}"))
                { borderParams["color_bgr"] = kv.Value; break; }
            }
            var mBW = Regex.Match(instruction, @"\u8fb9\u6846\s*(?:\u52a0\u7c97|\u7c97\u7ec6|\u5bbd\u5ea6)\s*(?:\u6539[\u6210\u4e3a]?)?\s*(\d+(?:\.\d+)?)");
            if (mBW.Success) borderParams["weight"] = float.Parse(mBW.Groups[1].Value);
            else if (instruction.Contains("\u8fb9\u6846\u52a0\u7c97")) borderParams["weight"] = 3.0f;
            if (borderParams.Count > 0)
                intents.Add(new Intent { Action = "set_border", Slide = sn, Target = Cl(target), Params = borderParams });

            var mMv = Regex.Match(instruction, @"(?:\u79fb\u52a8|\u4f4d\u7f6e)\s*(?:\u5230|\u8c03\u5230)?\s*\(?\s*(\d+)\s*[,\uff0c]\s*(\d+)\s*\)?");
            if (mMv.Success)
                intents.Add(new Intent { Action = "move_shape", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["left"] = int.Parse(mMv.Groups[1].Value), ["top"] = int.Parse(mMv.Groups[2].Value) } });
            else if (instruction.Contains("\u79fb\u52a8\u5230\u5de6\u4e0a"))
                intents.Add(new Intent { Action = "move_shape", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["left"] = 0, ["top"] = 0 } });

            var mW = Regex.Match(instruction, @"\u5bbd\u5ea6\s*(?:\u6539[\u6210\u4e3a]?|\u8c03[\u6210\u4e3a]?|\u8bbe\u4e3a)?\s*(\d+)");
            if (mW.Success)
                intents.Add(new Intent { Action = "resize_shape", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["width"] = int.Parse(mW.Groups[1].Value) } });
            var mHt = Regex.Match(instruction, @"\u9ad8\u5ea6\s*(?:\u6539[\u6210\u4e3a]?|\u8c03[\u6210\u4e3a]?|\u8bbe\u4e3a)?\s*(\d+)");
            if (mHt.Success)
                intents.Add(new Intent { Action = "resize_shape", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["height"] = int.Parse(mHt.Groups[1].Value) } });
            if (instruction.Contains("\u653e\u5927") && !intents.Any(i => i.Action == "resize_shape"))
                intents.Add(new Intent { Action = "resize_shape", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["scale_factor"] = 1.5f } });
            else if (instruction.Contains("\u7f29\u5c0f") && !intents.Any(i => i.Action == "resize_shape"))
                intents.Add(new Intent { Action = "resize_shape", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["scale_factor"] = 0.7f } });

            if (Regex.IsMatch(instruction, @"(?:\u5220\u9664|\u6e05\u9664|\u53bb\u6389)\s*(?:\u6240\u6709)?\s*\u52a8\u753b"))
                intents.Add(new Intent { Action = "remove_animation", Slide = sn, Target = Cl(target) });

            var mTxt = Regex.Match(instruction, @"(?:\u6539|\u6362|\u66ff\u6362|\u53d8)[\u6210\u4e3a]?\s*[\u201c\u201d\u300c](.+?)[\u201c\u201d\u300d]");
            if (mTxt.Success && !intents.Any(i => i.Action == "modify_cell" || i.Action == "add_textbox"))
                intents.Add(new Intent { Action = "modify_text", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["new_text"] = mTxt.Groups[1].Value } });

            var mFS = Regex.Match(instruction, @"\u5b57\u53f7?\s*(?:\u6539[\u6210\u4e3a]?|\u8c03[\u6210\u4e3a]?|\u8bbe\u4e3a)?\s*(\d+)");
            if (mFS.Success)
                intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["font_size"] = int.Parse(mFS.Groups[1].Value) } });
            else if (instruction.Contains("\u5927\u4e00\u70b9"))
                intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["font_size_factor"] = 1.3f } });
            else if (instruction.Contains("\u5c0f\u4e00\u70b9"))
                intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["font_size_factor"] = 0.75f } });

            if (CA(instruction, "\u52a0\u7c97", "\u7c97\u4f53", "bold"))
                intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["bold"] = true } });
            if (CA(instruction, "\u4e0b\u5212\u7ebf", "underline"))
                intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["underline"] = true } });
            if (CA(instruction, "\u5220\u9664\u7ebf", "\u5220\u9664\u7ebf\u6548\u679c", "strikethrough"))
                intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["strikethrough"] = true } });

            if (!intents.Any(i => i.Action == "set_fill" || i.Action == "set_border"))
            {
                foreach (var kv in ColorMap.Map)
                {
                    if (instruction.Contains(kv.Key) && Regex.IsMatch(instruction, $@"(?:\u6539|\u6362|\u53d8|\u8c03)[\u6210\u4e3a]?\s*{Regex.Escape(kv.Key)}"))
                    { intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target), Params = new Dictionary<string, object> { ["color"] = kv.Value } }); break; }
                }
            }

            if (CA(instruction, "\u5220\u9664", "\u5220\u6389", "\u53bb\u6389") &&
                !intents.Any(i => i.Action == "delete_slide" || i.Action == "remove_animation" || i.Action == "table_row_delete" || i.Action == "table_col_delete"))
                intents.Add(new Intent { Action = "delete", Slide = sn, Target = Cl(target) });

            var mAnim = Regex.Match(instruction, @"(?:\u6dfb\u52a0|\u52a0|\u8bbe\u7f6e)\s*\u52a8\u753b\s*(\S+)?");
            if (mAnim.Success)
            {
                string ef = mAnim.Groups[1].Success ? mAnim.Groups[1].Value : "appear";
                var cnMap = new Dictionary<string, string> { {"\u6de1\u5165","fade"},{"\u98de\u5165","fly"},{"\u51fa\u73b0","appear"},{"\u7f29\u653e","zoom"} };
                intents.Add(new Intent { Action = "animation", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["effect"] = cnMap.ContainsKey(ef) ? cnMap[ef] : ef } });
            }

            var mTr = Regex.Match(instruction, @"\u5207\u6362\s*(?:\u6548\u679c)?\s*(\S+)?");
            if (mTr.Success)
            {
                string tr = mTr.Groups[1].Success ? mTr.Groups[1].Value : "fade";
                var cnMap = new Dictionary<string, string> { {"\u6de1\u5316","fade"},{"\u63a8\u5165","push"},{"\u64e6\u9664","wipe"} };
                intents.Add(new Intent { Action = "transition", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["transition"] = cnMap.ContainsKey(tr) ? cnMap[tr] : tr } });
            }

            if (CA(instruction, "\u5bfc\u51fapdf", "\u5bfc\u51faPDF", "\u8f6cpdf", "\u8f6cPDF"))
                intents.Add(new Intent { Action = "export_pdf", Slide = sn, Target = Cl(target) });

            var mFN = Regex.Match(instruction, @"\u5b57\u4f53\s*(?:\u6539[\u6210\u4e3a]?|\u6362[\u6210\u4e3a]?|\u7528)?\s*(\S+)");
            if (mFN.Success && !CA(mFN.Groups[1].Value, "\u5927", "\u5c0f", "\u4e00\u70b9"))
                intents.Add(new Intent { Action = "modify_font", Slide = sn, Target = Cl(target),
                    Params = new Dictionary<string, object> { ["font_name"] = mFN.Groups[1].Value } });

            if (intents.Count == 0)
                intents.Add(new Intent { Action = "unknown", Raw = instruction, Target = Cl(target), Slide = sn });
            return intents;
        }
    }

    // ========== Runner ==========
    static class Runner
    {
        public static List<string> Run(PowerPointCOM ppt, List<Intent> intents, string output)
        {
            var changes = new List<string>();
            foreach (var intent in intents)
            {
                string a = intent.Action; var t = intent.Target; var p = intent.Params; int? sn = intent.Slide;
                if (a == "unknown") { Console.WriteLine($"\u26a0\ufe0f \u65e0\u6cd5\u7406\u89e3: {intent.Raw}"); continue; }
                if (a == "export_pdf") { changes.Add(ppt.ExportPdf(output.Replace(".pptx", ".pdf"))); continue; }
                if (a == "transition")
                {
                    var rng = sn.HasValue ? new List<int> { sn.Value } : Enumerable.Range(1, ppt.SlideCount).ToList();
                    foreach (int si in rng) changes.Add(ppt.SetTransition(si, p.ContainsKey("transition") ? (string)p["transition"] : "fade"));
                    continue;
                }
                if (a == "add_slide") { changes.Add($"\u2705 {ppt.AddSlide(p.ContainsKey("index") && p["index"] != null ? (int?)Convert.ToInt32(p["index"]) : null, p.ContainsKey("layout") ? Convert.ToInt32(p["layout"]) : 1)}"); continue; }
                if (a == "delete_slide") { changes.Add($"\u2705 {ppt.DeleteSlide(sn.Value)}"); continue; }
                if (a == "move_slide") { changes.Add($"\u2705 {ppt.MoveSlide(sn.Value, Convert.ToInt32(p["new_pos"]))}"); continue; }
                if (a == "add_textbox") { int si = sn ?? 1; changes.Add($"\u2705 {ppt.AddTextbox(si, (string)p["text"])}"); continue; }
                if (a == "add_picture") { int si = sn ?? 1; changes.Add($"\u2705 {ppt.AddPicture(si, (string)p["pic_path"])}"); continue; }
                if (a == "modify_cell") { int si = sn ?? 1; changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.ModifyCell(si, t, Convert.ToInt32(p["row"]), Convert.ToInt32(p["col"]), (string)p["text"])}"); continue; }
                if (a == "remove_animation")
                {
                    var rng = sn.HasValue ? new List<int> { sn.Value } : Enumerable.Range(1, ppt.SlideCount).ToList();
                    foreach (int si in rng)
                    {
                        try { changes.Add($"\u2705 {ppt.RemoveAnimation(si, p.ContainsKey("anim_index") ? (int?)Convert.ToInt32(p["anim_index"]) : null)}"); }
                        catch (Exception ex) { Console.WriteLine($"\u26a0\ufe0f \u7b2c{si}\u9875\u52a8\u753b\u64cd\u4f5c\u9519\u8bef: {ex.Message}"); }
                    }
                    continue;
                }

                var range = sn.HasValue ? new List<int> { sn.Value } : Enumerable.Range(1, ppt.SlideCount).ToList();
                foreach (int si in range)
                {
                    var shapes = ppt.FindShape(si, t);
                    if (shapes.Count == 0) { if (sn.HasValue) Console.WriteLine($"\u26a0\ufe0f \u7b2c{si}\u9875: \u6ca1\u627e\u5230"); continue; }
                    if (shapes.Count > 1) Console.WriteLine($"\u2139\ufe0f \u7b2c{si}\u9875: {shapes.Count} \u4e2a\u5339\u914d\uff0c\u5168\u90e8\u6267\u884c");
                    foreach (var shape in shapes)
                    {
                        try
                        {
                            if (a == "modify_text") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.ModifyText(shape, (string)p["new_text"])}");
                            else if (a == "modify_font") changes.Add($"\u2705 \u7b2c{si}\u9875 [{shape.Name}] {ppt.ModifyFont(shape, p)}");
                            else if (a == "delete") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.DeleteShape(shape)}");
                            else if (a == "animation") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.AddAnimation(si, shape, p.ContainsKey("effect") ? (string)p["effect"] : "appear")}");
                            else if (a == "set_alignment") changes.Add($"\u2705 \u7b2c{si}\u9875 [{shape.Name}] {ppt.SetAlignment(shape, (string)p["align"])}");
                            else if (a == "set_fill") changes.Add($"\u2705 \u7b2c{si}\u9875 [{shape.Name}] {ppt.SetFill(shape, Convert.ToInt32(p["color_bgr"]))}");
                            else if (a == "set_border") changes.Add($"\u2705 \u7b2c{si}\u9875 [{shape.Name}] {ppt.SetBorder(shape, p.ContainsKey("color_bgr") ? (int?)Convert.ToInt32(p["color_bgr"]) : null, p.ContainsKey("weight") ? (float?)Convert.ToSingle(p["weight"]) : null)}");
                            else if (a == "move_shape") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.MoveShape(shape, p.ContainsKey("left") ? (float?)Convert.ToSingle(p["left"]) : null, p.ContainsKey("top") ? (float?)Convert.ToSingle(p["top"]) : null)}");
                            else if (a == "resize_shape")
                            {
                                if (p.ContainsKey("scale_factor")) { float sf = Convert.ToSingle(p["scale_factor"]); changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.ResizeShape(shape, shape.Width * sf, shape.Height * sf)}"); }
                                else changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.ResizeShape(shape, p.ContainsKey("width") ? (float?)Convert.ToSingle(p["width"]) : null, p.ContainsKey("height") ? (float?)Convert.ToSingle(p["height"]) : null)}");
                            }
                            else if (a == "table_row_add") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.AddTableRow(shape)}");
                            else if (a == "table_row_delete") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.DeleteTableRow(shape, Convert.ToInt32(p["row"]))}");
                            else if (a == "table_col_add") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.AddTableColumn(shape)}");
                            else if (a == "table_col_delete") changes.Add($"\u2705 \u7b2c{si}\u9875 {ppt.DeleteTableColumn(shape, Convert.ToInt32(p["col"]))}");
                        }
                        catch (Exception ex) { Console.WriteLine($"\u26a0\ufe0f \u7b2c{si}\u9875 [{shape.Name}] \u9519\u8bef: {ex.Message}"); }
                    }
                }
            }

            if (changes.Count > 0)
            {
                ppt.Save(output);
                Console.WriteLine($"\n==================================================\n\U0001f4dd {changes.Count} \u9879\u4fee\u6539:");
                foreach (var c in changes) Console.WriteLine($"   {c}");
            }
            else Console.WriteLine("\n\u26a0\ufe0f \u6ca1\u6709\u6267\u884c\u4fee\u6539");
            return changes;
        }
    }

    // ========== Interactive ==========
    static class Interactive
    {
        public static void Start(PowerPointCOM ppt, string filepath)
        {
            Console.WriteLine("\n\U0001f3ae \u4ea4\u4e92\u6a21\u5f0f \u2014 \u8f93\u5165\u81ea\u7136\u8bed\u8a00\u6307\u4ee4\uff0c\u8f93\u5165 q \u9000\u51fa");
            Console.WriteLine("   COM \u72ec\u6709: \u52a8\u753b/\u5207\u6362\u6548\u679c/\u5bfc\u51faPDF/\u5bfc\u51fa\u56fe\u7247");
            Console.WriteLine("   \u65b0\u589e: \u6587\u672c\u6846/\u56fe\u7247/\u5e7b\u706f\u7247\u7ba1\u7406/\u8868\u683c\u64cd\u4f5c/\u5bf9\u9f50/\u586b\u5145/\u8fb9\u6846/\u79fb\u52a8/\u7f29\u653e");
            var desc = ppt.Inspect(); ppt.PrintStructure(desc);
            string output = filepath.Replace(".pptx", "_modified.pptx");
            while (true)
            {
                Console.Write("\n\U0001f4dd \u6307\u4ee4> ");
                string cmd = Console.ReadLine();
                if (cmd == null) break;
                cmd = cmd.Trim();
                if (cmd.ToLower() == "q" || cmd.ToLower() == "quit" || cmd.ToLower() == "exit") break;
                if (cmd == "inspect") { desc = ppt.Inspect(); ppt.PrintStructure(desc); continue; }
                if (cmd.StartsWith("export-image"))
                {
                    var mx = Regex.Match(cmd, @"(\d+)");
                    int si = mx.Success ? int.Parse(mx.Groups[1].Value) : 1;
                    ppt.ExportImage(si, $"slide_{si}.png"); continue;
                }
                var intents = IntentParser.ParseIntent(cmd);
                Console.WriteLine($"\U0001f50d {intents.Count} \u4e2a\u610f\u56fe: [{string.Join(", ", intents.Select(i => i.Action))}]");
                Runner.Run(ppt, intents, output);
            }
        }
    }

    // ========== Main ==========
    class Program
    {
        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;
            if (args.Length < 1)
            {
                Console.WriteLine("PPTX COM \u7f16\u8f91\u5668 (Windows + Office) \u2014 C# \u7248");
                Console.WriteLine("用法: pptx_editor_com.exe <file.pptx> [\u6307\u4ee4]");
                Console.WriteLine("      pptx_editor_com.exe <file.pptx> --inspect");
                Console.WriteLine("      pptx_editor_com.exe <file.pptx> --interactive");
                Console.WriteLine("      pptx_editor_com.exe <file.pptx> --export-images");
                return;
            }

            string path = args[0];
            if (!File.Exists(path)) { Console.WriteLine($"\u274c \u6587\u4ef6\u4e0d\u5b58\u5728: {path}"); return; }

            using (var ppt = new PowerPointCOM(false))
            {
                ppt.Open(path);

                if (args.Contains("--inspect"))
                {
                    var desc = ppt.Inspect(); ppt.PrintStructure(desc);
                    Console.WriteLine($"\n\U0001f4c4 \u7ed3\u6784\u5df2\u8f93\u51fa");
                }
                else if (args.Contains("--interactive"))
                {
                    Interactive.Start(ppt, path);
                }
                else if (args.Contains("--export-images"))
                {
                    for (int si = 1; si <= ppt.SlideCount; si++)
                        ppt.ExportImage(si, $"slide_{si}.png");
                }
                else if (args.Length > 1 && !args[1].StartsWith("--"))
                {
                    var argList = args.Skip(1).ToList();
                    string output = null;
                    int oi = argList.IndexOf("--output");
                    if (oi >= 0 && oi + 1 < argList.Count) { output = argList[oi + 1]; argList.RemoveAt(oi + 1); argList.RemoveAt(oi); }
                    string instruction = string.Join(" ", argList);
                    Console.WriteLine($"\U0001f4dd \u6307\u4ee4: {instruction}");
                    var intents = IntentParser.ParseIntent(instruction);
                    Console.WriteLine($"\U0001f50d {intents.Count} \u4e2a\u610f\u56fe: [{string.Join(", ", intents.Select(i => i.Action))}]");
                    if (output == null) output = path.Replace(".pptx", "_modified.pptx");
                    Runner.Run(ppt, intents, output);
                }
                else
                {
                    Console.WriteLine("\u26a0\ufe0f \u8bf7\u63d0\u4f9b\u6307\u4ee4\u6216\u4f7f\u7528 --inspect / --interactive / --export-images");
                }
            }
        }
    }
}
