using System.Globalization;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace PptInteropHost;

/// <summary>
/// Persistent PowerPoint Interop host. Mirrors the VBA bridge protocol so the
/// Python benchmark harness can drive it the same way it drives pywin32 / VBA.
///
/// Protocol: newline-delimited JSON over stdin/stdout.
///   Request : {"cmd":"open","path":"...","visible":true}
///   Response: {"ok":true,"result":<value>}  | {"ok":false,"error":"..."}
///
/// COM is accessed via late binding (dynamic over IDispatch) — exactly the same
/// mechanism pywin32 uses — so the benchmark compares the COM round-trip cost,
/// not early-bound PIA overhead. Zero NuGet / PIA / GUID dependency.
/// </summary>
internal static class Program
{
    // MsoTriState
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // Shape.Type
    private const int MsoPicture = 13;
    private const int MsoMedia = 16;
    private const int MsoTextBox = 17;

    // SaveAs formats
    private const int PpSaveAsPresentation = 1;          // .ppt
    private const int PpSaveAsOpenXmlPresentation = 24;  // .pptx

    private static readonly JsonSerializerOptions JsonOut = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    private static dynamic _app;
    private static dynamic _prs;

    private static int Main()
    {
        Console.InputEncoding = Encoding.UTF8;
        Console.OutputEncoding = Encoding.UTF8;

        string line;
        while ((line = Console.In.ReadLine()) != null)
        {
            if (line.Length == 0)
            {
                continue;
            }

            string responseJson;
            bool quit = false;
            try
            {
                using JsonDocument doc = JsonDocument.Parse(line);
                JsonElement req = doc.RootElement;
                string cmd = GetStr(req, "cmd", "");

                object result = Dispatch(cmd, req, out quit);
                responseJson = JsonSerializer.Serialize(
                    new Dictionary<string, object> { ["ok"] = true, ["result"] = result }, JsonOut);
            }
            catch (Exception ex)
            {
                responseJson = JsonSerializer.Serialize(
                    new Dictionary<string, object> { ["ok"] = false, ["error"] = FlattenMessage(ex) }, JsonOut);
            }

            Console.Out.WriteLine(responseJson);
            Console.Out.Flush();

            if (quit)
            {
                break;
            }
        }

        return 0;
    }

    private static object Dispatch(string cmd, JsonElement req, out bool quit)
    {
        quit = false;
        switch (cmd)
        {
            case "ping":
                return "pong";

            case "open":
                Open(GetStr(req, "path", ""), GetBool(req, "visible", false));
                return SlideCount();

            case "save":
                Save(GetStr(req, "path", null));
                return "saved";

            case "inspect":
                return InspectPresentation();

            case "inspect_slide":
                return InspectSlide(GetInt(req, "slide", 1));

            case "execute_action":
                return ExecuteAction(req.GetProperty("action"));

            case "set_notes":
                return SetNotes(GetInt(req, "slide", 1), GetStr(req, "text", ""));

            case "append_notes":
                return AppendNotes(GetInt(req, "slide", 1), GetStr(req, "text", ""), GetStr(req, "separator", "\n"));

            case "close":
            case "quit":
                quit = true;
                Shutdown();
                return "bye";

            default:
                throw new InvalidOperationException($"未知命令: {cmd}");
        }
    }

    // ---------- presentation lifecycle ----------

    private static void Open(string path, bool visible)
    {
        if (string.IsNullOrEmpty(path))
        {
            throw new ArgumentException("缺少 path");
        }

        if (_app == null)
        {
            Type t = Type.GetTypeFromProgID("PowerPoint.Application")
                     ?? throw new InvalidOperationException("无法创建 PowerPoint.Application (Office 未安装?)");
            _app = Activator.CreateInstance(t);
        }

        // PowerPoint forbids Application.Visible = msoFalse ("Hiding the application
        // window is not allowed"). Only force it visible; otherwise open the
        // presentation windowless via the WithWindow flag.
        if (visible)
        {
            _app.Visible = MsoTrue;
        }

        // Open(FileName, ReadOnly, Untitled, WithWindow)
        _prs = _app.Presentations.Open(path, MsoFalse, MsoFalse, visible ? MsoTrue : MsoFalse);
    }

    private static void Save(string path)
    {
        if (string.IsNullOrEmpty(path))
        {
            _prs.Save();
            return;
        }

        int fmt = Path.GetExtension(path).ToLowerInvariant() switch
        {
            ".pptx" => PpSaveAsOpenXmlPresentation,
            ".ppt" => PpSaveAsPresentation,
            _ => PpSaveAsOpenXmlPresentation,
        };
        _prs.SaveAs(path, fmt);
    }

    private static void Shutdown()
    {
        try
        {
            if (_prs != null)
            {
                _prs.Saved = MsoTrue;
                _prs.Close();
            }
        }
        catch
        {
            // ignore
        }

        try
        {
            _app?.Quit();
        }
        catch
        {
            // ignore
        }

        _prs = null;
        _app = null;
    }

    private static int SlideCount()
    {
        return (int)_prs.Slides.Count;
    }

    // ---------- inspect ----------

    private static object InspectPresentation()
    {
        var slides = new List<object>();
        int count = (int)_prs.Slides.Count;
        for (int si = 1; si <= count; si++)
        {
            slides.Add(BuildSlide(_prs.Slides[si], si));
        }

        return new Dictionary<string, object> { ["slides"] = slides };
    }

    private static object InspectSlide(int slideIndex)
    {
        var slides = new List<object> { BuildSlide(_prs.Slides[slideIndex], slideIndex) };
        return new Dictionary<string, object> { ["slides"] = slides };
    }

    private static Dictionary<string, object> BuildSlide(dynamic slide, int index)
    {
        var sd = new Dictionary<string, object>
        {
            ["index"] = index,
            ["layout"] = SlideLayoutName(slide),
        };

        double sw = (double)_prs.PageSetup.SlideWidth;
        double sh = (double)_prs.PageSetup.SlideHeight;

        var elements = new List<object>();
        int shapeCount = (int)slide.Shapes.Count;
        // Iterate by index (1-based) to avoid COM enumerator failures on large decks.
        for (int i = 1; i <= shapeCount; i++)
        {
            elements.Add(BuildElement(slide.Shapes[i], sw, sh));
        }

        sd["elements"] = elements;
        return sd;
    }

    private static Dictionary<string, object> BuildElement(dynamic shp, double sw, double sh)
    {
        double left = (double)shp.Left;
        double top = (double)shp.Top;
        double width = (double)shp.Width;
        double height = (double)shp.Height;
        double cx = left + width / 2.0;
        double cy = top + height / 2.0;

        var e = new Dictionary<string, object>
        {
            ["id"] = (int)shp.Id,
            ["name"] = (string)shp.Name,
            ["type"] = (int)shp.Type,
            ["left"] = Math.Round(left, 1),
            ["top"] = Math.Round(top, 1),
            ["width"] = Math.Round(width, 1),
            ["height"] = Math.Round(height, 1),
            ["text"] = ShapeText(shp),
            ["is_placeholder"] = false,
            ["has_image"] = false,
            ["has_chart"] = false,
            ["has_table"] = false,
            ["has_media"] = false,
            ["position_label"] = PositionLabel(cx, cy, sw, sh),
        };

        DetectContent(shp, e);

        try
        {
            dynamic pf = shp.PlaceholderFormat;
            if (pf != null)
            {
                int phType = (int)pf.Type;
                e["is_placeholder"] = true;
                e["ph_type"] = phType;
                e["ph_type_name"] = PlaceholderTypeName(phType);
            }
        }
        catch
        {
            // not a placeholder
        }

        return e;
    }

    private static void DetectContent(dynamic shp, Dictionary<string, object> e)
    {
        try
        {
            if (IsTrue(shp.HasTable))
            {
                e["has_table"] = true;
            }
        }
        catch
        {
        }

        try
        {
            if (IsTrue(shp.HasChart))
            {
                e["has_chart"] = true;
            }
        }
        catch
        {
        }

        try
        {
            int t = (int)shp.Type;
            if (t == MsoPicture)
            {
                e["has_image"] = true;
            }

            if (t == MsoMedia)
            {
                e["has_media"] = true;
            }
        }
        catch
        {
        }
    }

    // ---------- action dispatch ----------

    private static string ExecuteAction(JsonElement action)
    {
        string name = GetStr(action, "action", "");
        switch (name)
        {
            case "modify_text":
                return ModifyText(action);
            case "modify_font":
                return ModifyFont(action);
            case "set_alignment":
                return SetAlignment(action);
            case "set_fill":
                return SetFill(action);
            case "set_border":
                return SetBorder(action);
            case "move_shape":
                return MoveShape(action);
            case "resize_shape":
                return ResizeShape(action);
            case "set_zorder":
                return SetZOrder(action);
            case "delete":
            case "delete_shape":
                return DeleteShape(action);
            case "add_textbox":
                return AddTextbox(action);
            case "set_slide_background":
                return SetSlideBackground(action);
            case "add_slide":
                return AddSlide(action);
            case "delete_slide":
                return DeleteSlide(action);
            case "duplicate_slide":
                return DuplicateSlide(action);
            case "modify_cell":
                return ModifyCell(action);
            case "transition":
                return SetTransition(action);
            case "set_notes":
                return SetNotes(GetActionSlide(action), GetParamStr(action, "text", ""));
            case "append_notes":
                return AppendNotes(GetActionSlide(action), GetParamStr(action, "text", ""), GetParamStr(action, "separator", "\n"));
            case "sleep":
                System.Threading.Thread.Sleep((int)(GetParamDouble(action, "seconds", 0) * 1000));
                return "slept";
            default:
                throw new InvalidOperationException($"未支持的 action: {name}");
        }
    }

    private static string ModifyText(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        string text = GetParamStr(action, "text", "");
        shp.TextFrame.TextRange.Text = text;
        return $"文本 → '{Truncate(text, 20)}'";
    }

    private static string ModifyFont(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        if (!HasText(shp))
        {
            return $"跳过 [{(string)shp.Name}] (无文本框)";
        }

        dynamic font = shp.TextFrame.TextRange.Font;
        var changes = new List<string>();

        if (TryGetParamDouble(action, "font_size", out double size))
        {
            font.Size = size;
            changes.Add($"字号→{size}");
        }

        if (TryGetParamDouble(action, "font_size_factor", out double factor))
        {
            double old = (double)font.Size;
            if (old > 0)
            {
                double nv = Math.Round(old * factor, 1);
                font.Size = nv;
                changes.Add($"字号 {old}→{nv}");
            }
        }

        if (TryGetParamBool(action, "bold", out bool bold))
        {
            font.Bold = bold ? MsoTrue : MsoFalse;
            changes.Add(bold ? "加粗" : "取消加粗");
        }

        if (TryGetParamBool(action, "italic", out bool italic))
        {
            font.Italic = italic ? MsoTrue : MsoFalse;
            changes.Add("斜体");
        }

        if (TryGetParamBool(action, "underline", out bool underline))
        {
            font.Underline = underline ? MsoTrue : MsoFalse;
            changes.Add(underline ? "下划线" : "取消下划线");
        }

        if (TryGetParamInt(action, "color", out int color))
        {
            font.Color.RGB = color;
            changes.Add($"颜色→0x{color:X}");
        }

        if (TryGetParamStr(action, "font_name", out string fontName))
        {
            font.Name = fontName;
            changes.Add($"字体→{fontName}");
        }

        return string.Join(", ", changes);
    }

    private static string SetAlignment(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        if (!HasText(shp))
        {
            return $"跳过 [{(string)shp.Name}] (无文本框)";
        }

        string align = GetParamStr(action, "align", "left");
        int val = align.ToLowerInvariant() switch
        {
            "左" or "left" => 1,
            "居中" or "center" => 2,
            "右" or "right" => 3,
            "两端" or "justify" => 4,
            _ => 1,
        };

        dynamic tr = shp.TextFrame.TextRange;
        int pcount = (int)tr.Paragraphs().Count;
        for (int pi = 1; pi <= pcount; pi++)
        {
            tr.Paragraphs(pi).ParagraphFormat.Alignment = val;
        }

        return $"对齐 → {align}";
    }

    private static string SetFill(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        int color = GetParamInt(action, "color", 0);
        shp.Fill.Solid();
        shp.Fill.ForeColor.RGB = color;
        return $"填充 → 0x{color:X}";
    }

    private static string SetBorder(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        var changes = new List<string>();
        if (TryGetParamInt(action, "color", out int color))
        {
            shp.Line.ForeColor.RGB = color;
            changes.Add($"边框色→0x{color:X}");
        }

        if (TryGetParamDouble(action, "weight", out double weight))
        {
            shp.Line.Weight = weight;
            changes.Add($"边框粗→{weight}");
        }

        return changes.Count > 0 ? string.Join(", ", changes) : "边框未修改";
    }

    private static string MoveShape(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        var changes = new List<string>();
        if (TryGetParamDouble(action, "left", out double left))
        {
            shp.Left = left;
            changes.Add($"Left→{left}");
        }

        if (TryGetParamDouble(action, "top", out double top))
        {
            shp.Top = top;
            changes.Add($"Top→{top}");
        }

        return $"移动 {string.Join(", ", changes)}";
    }

    private static string ResizeShape(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        var changes = new List<string>();
        if (TryGetParamDouble(action, "width", out double width))
        {
            shp.Width = width;
            changes.Add($"Width→{width}");
        }

        if (TryGetParamDouble(action, "height", out double height))
        {
            shp.Height = height;
            changes.Add($"Height→{height}");
        }

        return $"缩放 {string.Join(", ", changes)}";
    }

    private static string SetZOrder(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        string order = GetParamStr(action, "order", "front");
        int cmd = order.ToLowerInvariant() switch
        {
            "front" => 0,
            "back" => 1,
            "forward" => 2,
            "backward" => 3,
            _ => 0,
        };
        shp.ZOrder(cmd);
        return $"层级 → {order}";
    }

    private static string DeleteShape(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        string n = (string)shp.Name;
        shp.Delete();
        return $"删除 [{n}]";
    }

    private static string AddTextbox(JsonElement action)
    {
        dynamic slide = _prs.Slides[GetActionSlide(action)];
        double left = GetParamDouble(action, "left", 100);
        double top = GetParamDouble(action, "top", 100);
        double width = GetParamDouble(action, "width", 200);
        double height = GetParamDouble(action, "height", 40);
        // AddTextbox(Orientation=msoTextOrientationHorizontal=1, Left, Top, Width, Height)
        dynamic box = slide.Shapes.AddTextbox(1, left, top, width, height);
        box.TextFrame.TextRange.Text = GetParamStr(action, "text", "");
        return $"新增文本框 [{(string)box.Name}]";
    }

    private static string SetSlideBackground(JsonElement action)
    {
        dynamic slide = _prs.Slides[GetActionSlide(action)];
        int color = GetParamInt(action, "color", 0xFFFFFF);
        slide.FollowMasterBackground = MsoFalse;
        slide.Background.Fill.Solid();
        slide.Background.Fill.ForeColor.RGB = color;
        return $"背景 → 0x{color:X}";
    }

    private static string AddSlide(JsonElement action)
    {
        int count = (int)_prs.Slides.Count;
        int index = GetParamInt(action, "index", count + 1);
        int layout = GetParamInt(action, "layout", 12); // ppLayoutBlank
        _prs.Slides.Add(index, layout);
        return $"新增幻灯片 @ {index}";
    }

    private static string DeleteSlide(JsonElement action)
    {
        int index = GetActionSlide(action);
        _prs.Slides[index].Delete();
        return $"删除幻灯片 {index}";
    }

    private static string DuplicateSlide(JsonElement action)
    {
        int index = GetActionSlide(action);
        _prs.Slides[index].Duplicate();
        return $"复制幻灯片 {index}";
    }

    private static string ModifyCell(JsonElement action)
    {
        dynamic shp = RequireShape(action);
        int row = GetParamInt(action, "row", 1);
        int col = GetParamInt(action, "col", 1);
        string text = GetParamStr(action, "text", "");
        shp.Table.Cell(row, col).Shape.TextFrame.TextRange.Text = text;
        return $"单元格 ({row},{col}) → '{Truncate(text, 20)}'";
    }

    private static string SetTransition(JsonElement action)
    {
        int index = GetActionSlide(action);
        string effect = GetParamStr(action, "transition", "fade");
        int entry = effect.ToLowerInvariant() switch
        {
            "fade" => 3849,
            "push" => 3334,
            "wipe" => 769,
            "split" => 3073,
            "dissolve" => 1537,
            "cut" => 257,
            "cover" => 1025,
            "uncover" => 1793,
            "random" => 513,
            "none" => 0,
            _ => 2745,
        };
        dynamic t = _prs.Slides[index].SlideShowTransition;
        t.EntryEffect = entry;
        if (TryGetParamDouble(action, "duration", out double duration))
        {
            t.Duration = duration;
        }

        return $"切换 → {effect}";
    }

    // ---------- notes ----------

    private static string SetNotes(int slideIndex, string text)
    {
        _prs.Slides[slideIndex].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = text;
        return "备注已更新";
    }

    private static string AppendNotes(int slideIndex, string text, string separator)
    {
        dynamic tr = _prs.Slides[slideIndex].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange;
        string current = (string)tr.Text;
        tr.Text = string.IsNullOrEmpty(current) ? text : current + separator + text;
        return "备注已追加";
    }

    // ---------- shape finding (mirrors pywin32 find_shape) ----------

    private static dynamic RequireShape(JsonElement action)
    {
        int slideIndex = GetActionSlide(action);
        JsonElement target = action.TryGetProperty("target", out JsonElement t) ? t : default;
        dynamic shp = FindFirstShape(slideIndex, target);
        if (shp == null)
        {
            throw new InvalidOperationException("未找到匹配的 shape");
        }

        return shp;
    }

    private static dynamic FindFirstShape(int slideIndex, JsonElement target)
    {
        dynamic slide = _prs.Slides[slideIndex];
        double sw = (double)_prs.PageSetup.SlideWidth;
        double sh = (double)_prs.PageSetup.SlideHeight;
        int count = (int)slide.Shapes.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic shp = slide.Shapes[i];
            if (MatchShape(shp, i, target, sw, sh))
            {
                return shp;
            }
        }

        return null;
    }

    private static bool MatchShape(dynamic shp, int shapeIndex, JsonElement target, double sw, double sh)
    {
        if (target.ValueKind != JsonValueKind.Object)
        {
            return true;
        }

        if (target.TryGetProperty("type", out JsonElement typeEl))
        {
            if (!MatchShapeType(shp, typeEl.GetString()?.ToLowerInvariant() ?? ""))
            {
                return false;
            }
        }

        if (target.TryGetProperty("text_match", out JsonElement tmEl))
        {
            string txt = ShapeText(shp);
            if (txt.IndexOf(tmEl.GetString() ?? "", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }
        }

        if (target.TryGetProperty("name", out JsonElement nameEl))
        {
            string nm = (string)shp.Name;
            if (nm.IndexOf(nameEl.GetString() ?? "", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }
        }

        if (target.TryGetProperty("position", out JsonElement posEl))
        {
            double cx = (double)shp.Left + (double)shp.Width / 2.0;
            double cy = (double)shp.Top + (double)shp.Height / 2.0;
            if (PositionLabel(cx, cy, sw, sh) != posEl.GetString())
            {
                return false;
            }
        }

        if (target.TryGetProperty("id", out JsonElement idEl) && idEl.ValueKind == JsonValueKind.Number)
        {
            if ((int)shp.Id != idEl.GetInt32())
            {
                return false;
            }
        }

        if (target.TryGetProperty("index", out JsonElement idxEl) && idxEl.ValueKind == JsonValueKind.Number)
        {
            if (shapeIndex != idxEl.GetInt32())
            {
                return false;
            }
        }

        return true;
    }

    private static bool MatchShapeType(dynamic shp, string targetType)
    {
        try
        {
            switch (targetType)
            {
                case "title":
                    int tt = (int)shp.PlaceholderFormat.Type;
                    return tt == 1 || tt == 3;
                case "subtitle":
                    return (int)shp.PlaceholderFormat.Type == 4;
                case "body":
                    int bt = (int)shp.PlaceholderFormat.Type;
                    return bt == 2 || bt == 7;
                case "picture":
                    return (int)shp.Type == MsoPicture;
                case "textbox":
                    return (int)shp.Type == MsoTextBox;
                case "chart":
                    return IsTrue(shp.HasChart);
                case "table":
                    return IsTrue(shp.HasTable);
                default:
                    return false;
            }
        }
        catch
        {
            return false;
        }
    }

    // ---------- helpers ----------

    private static bool HasText(dynamic shp)
    {
        try
        {
            return IsTrue(shp.HasTextFrame);
        }
        catch
        {
            return false;
        }
    }

    private static string ShapeText(dynamic shp)
    {
        try
        {
            if (IsTrue(shp.HasTextFrame))
            {
                return (string)shp.TextFrame.TextRange.Text;
            }
        }
        catch
        {
        }

        return "";
    }

    /// <summary>COM HasXxx properties return MsoTriState (-1/0), not bool.</summary>
    private static bool IsTrue(dynamic value)
    {
        try
        {
            return (int)value != 0;
        }
        catch
        {
            try
            {
                return (bool)value;
            }
            catch
            {
                return false;
            }
        }
    }

    private static string SlideLayoutName(dynamic slide)
    {
        try
        {
            return (string)slide.CustomLayout.Name;
        }
        catch
        {
            try
            {
                return ((int)slide.Layout).ToString(CultureInfo.InvariantCulture);
            }
            catch
            {
                return "";
            }
        }
    }

    private static string PositionLabel(double cx, double cy, double sw, double sh)
    {
        string h = cx < sw * 0.33 ? "左" : (cx > sw * 0.67 ? "右" : "中");
        string v = cy < sh * 0.33 ? "上" : (cy > sh * 0.67 ? "下" : "中");
        return h + v;
    }

    private static string PlaceholderTypeName(int phType)
    {
        return phType switch
        {
            1 => "TITLE",
            2 => "BODY",
            3 => "CENTER_TITLE",
            4 => "SUBTITLE",
            7 => "OBJECT",
            8 => "CHART",
            9 => "TABLE",
            12 => "MEDIA",
            13 => "SLIDE_NUMBER",
            15 => "FOOTER",
            _ => $"({phType})",
        };
    }

    private static string Truncate(string s, int n)
    {
        if (string.IsNullOrEmpty(s))
        {
            return s ?? "";
        }

        return s.Length <= n ? s : s.Substring(0, n);
    }

    private static string FlattenMessage(Exception ex)
    {
        return ex.Message.Replace("\r", " ").Replace("\n", " ");
    }

    // ---------- JSON accessors ----------

    private static int GetActionSlide(JsonElement action)
    {
        return GetInt(action, "slide", 1);
    }

    private static JsonElement Params(JsonElement action)
    {
        return action.TryGetProperty("params", out JsonElement p) ? p : default;
    }

    private static string GetStr(JsonElement obj, string name, string def)
    {
        if (obj.ValueKind == JsonValueKind.Object && obj.TryGetProperty(name, out JsonElement v) && v.ValueKind == JsonValueKind.String)
        {
            return v.GetString();
        }

        return def;
    }

    private static bool GetBool(JsonElement obj, string name, bool def)
    {
        if (obj.ValueKind == JsonValueKind.Object && obj.TryGetProperty(name, out JsonElement v))
        {
            if (v.ValueKind == JsonValueKind.True)
            {
                return true;
            }

            if (v.ValueKind == JsonValueKind.False)
            {
                return false;
            }
        }

        return def;
    }

    private static int GetInt(JsonElement obj, string name, int def)
    {
        if (obj.ValueKind == JsonValueKind.Object && obj.TryGetProperty(name, out JsonElement v) && v.ValueKind == JsonValueKind.Number)
        {
            return v.GetInt32();
        }

        return def;
    }

    private static string GetParamStr(JsonElement action, string name, string def)
    {
        return GetStr(Params(action), name, def);
    }

    private static bool TryGetParamStr(JsonElement action, string name, out string value)
    {
        JsonElement p = Params(action);
        if (p.ValueKind == JsonValueKind.Object && p.TryGetProperty(name, out JsonElement v) && v.ValueKind == JsonValueKind.String)
        {
            value = v.GetString();
            return true;
        }

        value = null;
        return false;
    }

    private static double GetParamDouble(JsonElement action, string name, double def)
    {
        return TryGetParamDouble(action, name, out double v) ? v : def;
    }

    private static bool TryGetParamDouble(JsonElement action, string name, out double value)
    {
        JsonElement p = Params(action);
        if (p.ValueKind == JsonValueKind.Object && p.TryGetProperty(name, out JsonElement v) && v.ValueKind == JsonValueKind.Number)
        {
            value = v.GetDouble();
            return true;
        }

        value = 0;
        return false;
    }

    private static int GetParamInt(JsonElement action, string name, int def)
    {
        return TryGetParamInt(action, name, out int v) ? v : def;
    }

    private static bool TryGetParamInt(JsonElement action, string name, out int value)
    {
        JsonElement p = Params(action);
        if (p.ValueKind == JsonValueKind.Object && p.TryGetProperty(name, out JsonElement v) && v.ValueKind == JsonValueKind.Number)
        {
            value = v.GetInt32();
            return true;
        }

        value = 0;
        return false;
    }

    private static bool TryGetParamBool(JsonElement action, string name, out bool value)
    {
        JsonElement p = Params(action);
        if (p.ValueKind == JsonValueKind.Object && p.TryGetProperty(name, out JsonElement v))
        {
            if (v.ValueKind == JsonValueKind.True)
            {
                value = true;
                return true;
            }

            if (v.ValueKind == JsonValueKind.False)
            {
                value = false;
                return true;
            }
        }

        value = false;
        return false;
    }
}
