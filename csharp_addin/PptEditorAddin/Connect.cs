using System;
using System.Collections.Generic;
using System.Globalization;
using System.Runtime.InteropServices;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PptEditorAddin
{
    // ---- ext_ConnectMode / ext_DisconnectMode (verbatim from the Office contract) ----

    [Guid("289E9AF1-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3,
        ext_cm_Solution = 4,
        ext_cm_UISetup = 5,
    }

    [Guid("289E9AF2-4973-11D1-AE81-00A0C90F26F4")]
    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1,
        ext_dm_UISetupComplete = 2,
        ext_dm_SolutionClosed = 3,
    }

    /// <summary>
    /// IDTExtensibility2 declared by hand (no Extensibility PIA dependency) so the
    /// project builds with just the .NET SDK. This is the battle-tested Excel-DNA
    /// definition: a DUAL interface whose 'custom' parameter is marshalled as a
    /// SAFEARRAY(VARIANT). Omitting the SafeArray attribute makes Office's
    /// OnConnection call fail during parameter marshalling before the body runs.
    /// </summary>
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] ext_ConnectMode ConnectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ext_DisconnectMode RemoveMode,
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown(
            [In, MarshalAs(UnmanagedType.SafeArray, SafeArraySubType = VarEnum.VT_VARIANT)] ref Array custom);
    }

    /// <summary>
    /// In-process PowerPoint COM add-in. Runs inside POWERPNT.EXE, so the PowerPoint
    /// object model is accessed in-process (zero cross-process COM marshalling) — the
    /// same reason VBA is far faster than out-of-process pywin32 / the C# Interop exe.
    ///
    /// External drivers (Python via pywin32) reach the bridge via
    /// Application.COMAddIns.Item("PptEditor.AddIn").Object and make ONE coarse
    /// cross-process call per operation (Ping / InspectJson / InspectSlideJson /
    /// ExecuteActionJson); all per-shape traversal happens in-process — mirroring VBA.
    ///
    /// ClassInterface.AutoDispatch exposes the public methods via a late-bound
    /// (IDispatch) class interface so pywin32 can call obj.InspectJson etc., while
    /// Office still resolves IDTExtensibility2 by its IID for lifecycle callbacks.
    /// </summary>
    [ComVisible(true)]
    [Guid("89E53E12-1EB0-4DDF-8017-16178D7DE66D")]
    [ProgId("PptEditor.AddIn")]
#pragma warning disable CS0618 // AutoDispatch is obsolete but is the correct choice for a late-bound add-in surface
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
#pragma warning restore CS0618
    public class Connect : IDTExtensibility2
    {
        // Shape.Type (Office.MsoShapeType) compared as int for brevity.
        private const int MsoPicture = 13;
        private const int MsoMedia = 16;
        private const int MsoTextBox = 17;

        // Early-bound PowerPoint Application (vtable calls, no IDispatch/DLR) => VBA-class speed.
        private PowerPoint.Application _app;

        private static void Log(string msg)
        {
            try
            {
                System.IO.File.AppendAllText(
                    System.IO.Path.Combine(System.IO.Path.GetTempPath(), "ppteditor_addin.log"),
                    DateTime.Now.ToString("HH:mm:ss.fff") + " " + msg + Environment.NewLine);
            }
            catch { }
        }

        // ---------- IDTExtensibility2 ----------

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            Log("OnConnection enter connectMode=" + connectMode);
            _app = (PowerPoint.Application)application;
            try
            {
                ((Office.COMAddIn)addInInst).Object = this;
                Log("OnConnection set addInInst.Object OK");
            }
            catch (Exception ex)
            {
                Log("OnConnection set Object FAILED: " + ex.Message);
            }
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            _app = null;
        }

        public void OnAddInsUpdate(ref Array custom) { }

        public void OnStartupComplete(ref Array custom) { }

        public void OnBeginShutdown(ref Array custom) { }

        public object RequestComAddInAutomationService()
        {
            return this;
        }

        // ---------- bridge surface ----------

        public string Ping()
        {
            return "pong";
        }

        public string InspectJson()
        {
            return JsonConvert.SerializeObject(InspectPresentation());
        }

        public string InspectSlideJson(int slideIndex)
        {
            return JsonConvert.SerializeObject(InspectSlide(slideIndex));
        }

        public string ExecuteActionJson(string actionJson)
        {
            JObject action = JObject.Parse(actionJson);
            // Accept both nested {"action":{...}} and flat {"action":"name",...}.
            if (action["action"] is JObject nested)
            {
                return ExecuteAction(nested);
            }

            return ExecuteAction(action);
        }

        // ---------- presentation resolution ----------

        private PowerPoint.Presentation Pres()
        {
            try
            {
                PowerPoint.Presentation p = _app.ActivePresentation;
                if (p != null)
                {
                    return p;
                }
            }
            catch
            {
            }

            return _app.Presentations[1];
        }

        // ---------- inspect ----------

        private object InspectPresentation()
        {
            PowerPoint.Presentation prs = Pres();
            var slides = new List<object>();
            int count = prs.Slides.Count;
            for (int si = 1; si <= count; si++)
            {
                slides.Add(BuildSlide(prs, prs.Slides[si], si));
            }

            return new Dictionary<string, object> { ["slides"] = slides };
        }

        private object InspectSlide(int slideIndex)
        {
            PowerPoint.Presentation prs = Pres();
            var slides = new List<object> { BuildSlide(prs, prs.Slides[slideIndex], slideIndex) };
            return new Dictionary<string, object> { ["slides"] = slides };
        }

        private Dictionary<string, object> BuildSlide(PowerPoint.Presentation prs, PowerPoint.Slide slide, int index)
        {
            var sd = new Dictionary<string, object>
            {
                ["index"] = index,
                ["layout"] = SlideLayoutName(slide),
            };

            double sw = prs.PageSetup.SlideWidth;
            double sh = prs.PageSetup.SlideHeight;

            var elements = new List<object>();
            int shapeCount = slide.Shapes.Count;
            for (int i = 1; i <= shapeCount; i++)
            {
                elements.Add(BuildElement(slide.Shapes[i], sw, sh));
            }

            sd["elements"] = elements;
            return sd;
        }

        private Dictionary<string, object> BuildElement(PowerPoint.Shape shp, double sw, double sh)
        {
            double left = shp.Left;
            double top = shp.Top;
            double width = shp.Width;
            double height = shp.Height;
            double cx = left + width / 2.0;
            double cy = top + height / 2.0;

            var e = new Dictionary<string, object>
            {
                ["id"] = shp.Id,
                ["name"] = shp.Name,
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
                PowerPoint.PlaceholderFormat pf = shp.PlaceholderFormat;
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

        private void DetectContent(PowerPoint.Shape shp, Dictionary<string, object> e)
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

        private string ExecuteAction(JObject action)
        {
            string name = (string)action["action"] ?? "";
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
                    return SetNotes(ActionSlide(action), ParStr(action, "text", ""));
                case "append_notes":
                    return AppendNotes(ActionSlide(action), ParStr(action, "text", ""), ParStr(action, "separator", "\n"));
                case "sleep":
                    System.Threading.Thread.Sleep((int)(ParDouble(action, "seconds", 0) * 1000));
                    return "slept";
                default:
                    throw new InvalidOperationException("未支持的 action: " + name);
            }
        }

        private string ModifyText(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            string text = ParStr(action, "text", "");
            shp.TextFrame.TextRange.Text = text;
            return "文本 → '" + Truncate(text, 20) + "'";
        }

        private string ModifyFont(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            if (!HasText(shp))
            {
                return "跳过 [" + shp.Name + "] (无文本框)";
            }

            PowerPoint.Font font = shp.TextFrame.TextRange.Font;
            var changes = new List<string>();

            if (TryParDouble(action, "font_size", out double size))
            {
                font.Size = (float)size;
                changes.Add("字号→" + size);
            }

            if (TryParDouble(action, "font_size_factor", out double factor))
            {
                double old = font.Size;
                if (old > 0)
                {
                    double nv = Math.Round(old * factor, 1);
                    font.Size = (float)nv;
                    changes.Add("字号 " + old + "→" + nv);
                }
            }

            if (TryParBool(action, "bold", out bool bold))
            {
                font.Bold = bold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                changes.Add(bold ? "加粗" : "取消加粗");
            }

            if (TryParBool(action, "italic", out bool italic))
            {
                font.Italic = italic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                changes.Add("斜体");
            }

            if (TryParBool(action, "underline", out bool underline))
            {
                font.Underline = underline ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse;
                changes.Add(underline ? "下划线" : "取消下划线");
            }

            if (TryParInt(action, "color", out int color))
            {
                font.Color.RGB = color;
                changes.Add("颜色→0x" + color.ToString("X"));
            }

            if (TryParStr(action, "font_name", out string fontName))
            {
                font.Name = fontName;
                changes.Add("字体→" + fontName);
            }

            return string.Join(", ", changes);
        }

        private string SetAlignment(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            if (!HasText(shp))
            {
                return "跳过 [" + shp.Name + "] (无文本框)";
            }

            string align = ParStr(action, "align", "left");
            PowerPoint.PpParagraphAlignment val;
            switch (align.ToLowerInvariant())
            {
                case "左":
                case "left":
                    val = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    break;
                case "居中":
                case "center":
                    val = PowerPoint.PpParagraphAlignment.ppAlignCenter;
                    break;
                case "右":
                case "right":
                    val = PowerPoint.PpParagraphAlignment.ppAlignRight;
                    break;
                case "两端":
                case "justify":
                    val = PowerPoint.PpParagraphAlignment.ppAlignJustify;
                    break;
                default:
                    val = PowerPoint.PpParagraphAlignment.ppAlignLeft;
                    break;
            }

            PowerPoint.TextRange tr = shp.TextFrame.TextRange;
            int pcount = tr.Paragraphs().Count;
            for (int pi = 1; pi <= pcount; pi++)
            {
                tr.Paragraphs(pi).ParagraphFormat.Alignment = val;
            }

            return "对齐 → " + align;
        }

        private string SetFill(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            int color = ParInt(action, "color", 0);
            shp.Fill.Solid();
            shp.Fill.ForeColor.RGB = color;
            return "填充 → 0x" + color.ToString("X");
        }

        private string SetBorder(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            var changes = new List<string>();
            if (TryParInt(action, "color", out int color))
            {
                shp.Line.ForeColor.RGB = color;
                changes.Add("边框色→0x" + color.ToString("X"));
            }

            if (TryParDouble(action, "weight", out double weight))
            {
                shp.Line.Weight = (float)weight;
                changes.Add("边框粗→" + weight);
            }

            return changes.Count > 0 ? string.Join(", ", changes) : "边框未修改";
        }

        private string MoveShape(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            var changes = new List<string>();
            if (TryParDouble(action, "left", out double left))
            {
                shp.Left = (float)left;
                changes.Add("Left→" + left);
            }

            if (TryParDouble(action, "top", out double top))
            {
                shp.Top = (float)top;
                changes.Add("Top→" + top);
            }

            return "移动 " + string.Join(", ", changes);
        }

        private string ResizeShape(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            var changes = new List<string>();
            if (TryParDouble(action, "width", out double width))
            {
                shp.Width = (float)width;
                changes.Add("Width→" + width);
            }

            if (TryParDouble(action, "height", out double height))
            {
                shp.Height = (float)height;
                changes.Add("Height→" + height);
            }

            return "缩放 " + string.Join(", ", changes);
        }

        private string SetZOrder(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            string order = ParStr(action, "order", "front");
            Office.MsoZOrderCmd cmd;
            switch (order.ToLowerInvariant())
            {
                case "front":
                    cmd = Office.MsoZOrderCmd.msoBringToFront;
                    break;
                case "back":
                    cmd = Office.MsoZOrderCmd.msoSendToBack;
                    break;
                case "forward":
                    cmd = Office.MsoZOrderCmd.msoBringForward;
                    break;
                case "backward":
                    cmd = Office.MsoZOrderCmd.msoSendBackward;
                    break;
                default:
                    cmd = Office.MsoZOrderCmd.msoBringToFront;
                    break;
            }

            shp.ZOrder(cmd);
            return "层级 → " + order;
        }

        private string DeleteShape(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            string n = shp.Name;
            shp.Delete();
            return "删除 [" + n + "]";
        }

        private string AddTextbox(JObject action)
        {
            PowerPoint.Slide slide = Pres().Slides[ActionSlide(action)];
            double left = ParDouble(action, "left", 100);
            double top = ParDouble(action, "top", 100);
            double width = ParDouble(action, "width", 200);
            double height = ParDouble(action, "height", 40);
            PowerPoint.Shape box = slide.Shapes.AddTextbox(
                Office.MsoTextOrientation.msoTextOrientationHorizontal,
                (float)left, (float)top, (float)width, (float)height);
            box.TextFrame.TextRange.Text = ParStr(action, "text", "");
            return "新增文本框 [" + box.Name + "]";
        }

        private string SetSlideBackground(JObject action)
        {
            PowerPoint.Slide slide = Pres().Slides[ActionSlide(action)];
            int color = ParInt(action, "color", 0xFFFFFF);
            slide.FollowMasterBackground = Office.MsoTriState.msoFalse;
            slide.Background.Fill.Solid();
            slide.Background.Fill.ForeColor.RGB = color;
            return "背景 → 0x" + color.ToString("X");
        }

        private string AddSlide(JObject action)
        {
            PowerPoint.Presentation prs = Pres();
            int count = prs.Slides.Count;
            int index = ParInt(action, "index", count + 1);
            int layout = ParInt(action, "layout", 12); // ppLayoutBlank
            prs.Slides.Add(index, (PowerPoint.PpSlideLayout)layout);
            return "新增幻灯片 @ " + index;
        }

        private string DeleteSlide(JObject action)
        {
            int index = ActionSlide(action);
            Pres().Slides[index].Delete();
            return "删除幻灯片 " + index;
        }

        private string DuplicateSlide(JObject action)
        {
            int index = ActionSlide(action);
            Pres().Slides[index].Duplicate();
            return "复制幻灯片 " + index;
        }

        private string ModifyCell(JObject action)
        {
            PowerPoint.Shape shp = RequireShape(action);
            int row = ParInt(action, "row", 1);
            int col = ParInt(action, "col", 1);
            string text = ParStr(action, "text", "");
            shp.Table.Cell(row, col).Shape.TextFrame.TextRange.Text = text;
            return "单元格 (" + row + "," + col + ") → '" + Truncate(text, 20) + "'";
        }

        private string SetTransition(JObject action)
        {
            int index = ActionSlide(action);
            string effect = ParStr(action, "transition", "fade");
            int entry;
            switch (effect.ToLowerInvariant())
            {
                case "fade":
                    entry = 3849;
                    break;
                case "push":
                    entry = 3334;
                    break;
                case "wipe":
                    entry = 769;
                    break;
                case "split":
                    entry = 3073;
                    break;
                case "dissolve":
                    entry = 1537;
                    break;
                case "cut":
                    entry = 257;
                    break;
                case "cover":
                    entry = 1025;
                    break;
                case "uncover":
                    entry = 1793;
                    break;
                case "random":
                    entry = 513;
                    break;
                case "none":
                    entry = 0;
                    break;
                default:
                    entry = 2745;
                    break;
            }

            PowerPoint.SlideShowTransition t = Pres().Slides[index].SlideShowTransition;
            t.EntryEffect = (PowerPoint.PpEntryEffect)entry;
            if (TryParDouble(action, "duration", out double duration))
            {
                t.Duration = (float)duration;
            }

            return "切换 → " + effect;
        }

        // ---------- notes ----------

        private string SetNotes(int slideIndex, string text)
        {
            Pres().Slides[slideIndex].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange.Text = text;
            return "备注已更新";
        }

        private string AppendNotes(int slideIndex, string text, string separator)
        {
            PowerPoint.TextRange tr = Pres().Slides[slideIndex].NotesPage.Shapes.Placeholders[2].TextFrame.TextRange;
            string current = tr.Text;
            tr.Text = string.IsNullOrEmpty(current) ? text : current + separator + text;
            return "备注已追加";
        }

        // ---------- shape finding (mirrors pywin32 find_shape) ----------

        private PowerPoint.Shape RequireShape(JObject action)
        {
            int slideIndex = ActionSlide(action);
            JToken target = action["target"];
            PowerPoint.Shape shp = FindFirstShape(slideIndex, target);
            if (shp == null)
            {
                throw new InvalidOperationException("未找到匹配的 shape");
            }

            return shp;
        }

        private PowerPoint.Shape FindFirstShape(int slideIndex, JToken target)
        {
            PowerPoint.Presentation prs = Pres();
            PowerPoint.Slide slide = prs.Slides[slideIndex];
            double sw = prs.PageSetup.SlideWidth;
            double sh = prs.PageSetup.SlideHeight;
            int count = slide.Shapes.Count;
            for (int i = 1; i <= count; i++)
            {
                PowerPoint.Shape shp = slide.Shapes[i];
                if (MatchShape(shp, i, target, sw, sh))
                {
                    return shp;
                }
            }

            return null;
        }

        private bool MatchShape(PowerPoint.Shape shp, int shapeIndex, JToken target, double sw, double sh)
        {
            if (!(target is JObject t))
            {
                return true;
            }

            JToken typeEl = t["type"];
            if (typeEl != null && typeEl.Type == JTokenType.String)
            {
                if (!MatchShapeType(shp, typeEl.ToString().ToLowerInvariant()))
                {
                    return false;
                }
            }

            JToken tmEl = t["text_match"];
            if (tmEl != null && tmEl.Type == JTokenType.String)
            {
                string txt = ShapeText(shp);
                if (txt.IndexOf(tmEl.ToString(), StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            JToken nameEl = t["name"];
            if (nameEl != null && nameEl.Type == JTokenType.String)
            {
                string nm = shp.Name;
                if (nm.IndexOf(nameEl.ToString(), StringComparison.OrdinalIgnoreCase) < 0)
                {
                    return false;
                }
            }

            JToken posEl = t["position"];
            if (posEl != null && posEl.Type == JTokenType.String)
            {
                double cx = shp.Left + shp.Width / 2.0;
                double cy = shp.Top + shp.Height / 2.0;
                if (PositionLabel(cx, cy, sw, sh) != posEl.ToString())
                {
                    return false;
                }
            }

            JToken idEl = t["id"];
            if (idEl != null && idEl.Type == JTokenType.Integer)
            {
                if (shp.Id != idEl.Value<int>())
                {
                    return false;
                }
            }

            JToken idxEl = t["index"];
            if (idxEl != null && idxEl.Type == JTokenType.Integer)
            {
                if (shapeIndex != idxEl.Value<int>())
                {
                    return false;
                }
            }

            return true;
        }

        private bool MatchShapeType(PowerPoint.Shape shp, string targetType)
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

        private static bool HasText(PowerPoint.Shape shp)
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

        private static string ShapeText(PowerPoint.Shape shp)
        {
            try
            {
                if (IsTrue(shp.HasTextFrame))
                {
                    return shp.TextFrame.TextRange.Text ?? "";
                }
            }
            catch
            {
            }

            return "";
        }

        /// <summary>COM HasXxx properties return MsoTriState (msoTrue = -1).</summary>
        private static bool IsTrue(Office.MsoTriState value)
        {
            return value == Office.MsoTriState.msoTrue;
        }

        private static string SlideLayoutName(PowerPoint.Slide slide)
        {
            try
            {
                return slide.CustomLayout.Name;
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
            switch (phType)
            {
                case 1: return "TITLE";
                case 2: return "BODY";
                case 3: return "CENTER_TITLE";
                case 4: return "SUBTITLE";
                case 7: return "OBJECT";
                case 8: return "CHART";
                case 9: return "TABLE";
                case 12: return "MEDIA";
                case 13: return "SLIDE_NUMBER";
                case 15: return "FOOTER";
                default: return "(" + phType + ")";
            }
        }

        private static string Truncate(string s, int n)
        {
            if (string.IsNullOrEmpty(s))
            {
                return s ?? "";
            }

            return s.Length <= n ? s : s.Substring(0, n);
        }

        // ---------- JSON param accessors ----------

        private static int ActionSlide(JObject a)
        {
            JToken t = a["slide"];
            return t != null && t.Type == JTokenType.Integer ? t.Value<int>() : 1;
        }

        private static JToken Par(JObject a, string name)
        {
            return (a["params"] as JObject)?[name];
        }

        private static string ParStr(JObject a, string name, string def)
        {
            JToken t = Par(a, name);
            return t == null || t.Type == JTokenType.Null ? def : t.ToString();
        }

        private static bool TryParStr(JObject a, string name, out string value)
        {
            JToken t = Par(a, name);
            if (t != null && t.Type == JTokenType.String)
            {
                value = t.ToString();
                return true;
            }

            value = null;
            return false;
        }

        private static bool TryParDouble(JObject a, string name, out double value)
        {
            JToken t = Par(a, name);
            if (t != null && (t.Type == JTokenType.Float || t.Type == JTokenType.Integer))
            {
                value = t.Value<double>();
                return true;
            }

            value = 0;
            return false;
        }

        private static double ParDouble(JObject a, string name, double def)
        {
            return TryParDouble(a, name, out double v) ? v : def;
        }

        private static bool TryParInt(JObject a, string name, out int value)
        {
            JToken t = Par(a, name);
            if (t != null && (t.Type == JTokenType.Integer || t.Type == JTokenType.Float))
            {
                value = t.Value<int>();
                return true;
            }

            value = 0;
            return false;
        }

        private static int ParInt(JObject a, string name, int def)
        {
            return TryParInt(a, name, out int v) ? v : def;
        }

        private static bool TryParBool(JObject a, string name, out bool value)
        {
            JToken t = Par(a, name);
            if (t != null && t.Type == JTokenType.Boolean)
            {
                value = t.Value<bool>();
                return true;
            }

            value = false;
            return false;
        }
    }
}
