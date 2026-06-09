using System.Collections.Generic;
using System.Text;

namespace PptInteropHost;

/// <summary>
/// CodeAct API surface exposed to a single <c>execute_code</c> run.
///
/// The model writes ONE C# script against these members; the script executes
/// once in-process (inside the PowerPoint COM host), collapsing what would
/// otherwise be many <c>execute_action</c> request/response round-trips into a
/// single in-process program. This mirrors the Agent Framework CodeAct pattern:
/// expose a single <c>execute_code</c> tool, let the model express the full plan
/// as a short program, and run it once instead of scattering it across many
/// tool-call turns.
///
/// Conventions (match COM_STANDARD.md):
///   - Colors are BGR, NOT RGB: red = 0x0000FF, blue = 0xFF0000.
///   - Positions / sizes are in points (72pt = 1 inch).
///   - All slide / shape indices are 1-based.
///
/// The script can also reach the raw late-bound COM objects via <see cref="App"/>
/// and <see cref="Prs"/> for anything not wrapped by a helper.
/// </summary>
public sealed class PptApi
{
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;
    private const int MsoTextOrientationHorizontal = 1;

    private readonly StringBuilder _out = new();

    /// <summary>Raw <c>PowerPoint.Application</c> COM object (late-bound).</summary>
    public dynamic App { get; }

    /// <summary>Raw active <c>Presentation</c> COM object (late-bound).</summary>
    public dynamic Prs { get; }

    public PptApi(dynamic app, dynamic prs)
    {
        App = app;
        Prs = prs;
    }

    /// <summary>Text captured by <see cref="Print"/>, returned to the caller as <c>result.output</c>.</summary>
    public string Output => _out.ToString();

    /// <summary>Append a line to the script output. This is CodeAct's <c>print()</c>.</summary>
    public void Print(object message) => _out.AppendLine(message?.ToString() ?? string.Empty);

    // ---------- navigation ----------

    /// <summary>Number of slides in the presentation.</summary>
    public int SlideCount => (int)Prs.Slides.Count;

    /// <summary>1-based slide accessor.</summary>
    public dynamic Slide(int slide) => Prs.Slides[slide];

    /// <summary>Shape by 1-based slide index and 1-based shape index.</summary>
    public dynamic Shape(int slide, int shapeIndex) => Prs.Slides[slide].Shapes[shapeIndex];

    /// <summary>Number of shapes on a slide.</summary>
    public int ShapeCount(int slide) => (int)Prs.Slides[slide].Shapes.Count;

    /// <summary>First shape on a slide whose text contains <paramref name="contains"/> (case-insensitive), or null.</summary>
    public dynamic FindByText(int slide, string contains)
    {
        dynamic shapes = Prs.Slides[slide].Shapes;
        int count = (int)shapes.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic shp = shapes[i];
            if (!HasText(shp))
            {
                continue;
            }

            string txt = (string)shp.TextFrame.TextRange.Text;
            if (!string.IsNullOrEmpty(txt) && txt.IndexOf(contains, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return shp;
            }
        }

        return null;
    }

    /// <summary>First title placeholder (type TITLE or CENTER_TITLE) on a slide, or null.</summary>
    public dynamic Title(int slide)
    {
        dynamic shapes = Prs.Slides[slide].Shapes;
        int count = (int)shapes.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic shp = shapes[i];
            try
            {
                int t = (int)shp.PlaceholderFormat.Type;
                if (t == 1 || t == 3)
                {
                    return shp;
                }
            }
            catch
            {
                // not a placeholder
            }
        }

        return null;
    }

    /// <summary>First shape on a slide whose Name equals <paramref name="name"/> (case-insensitive), or null.</summary>
    public dynamic FindByName(int slide, string name)
    {
        dynamic shapes = Prs.Slides[slide].Shapes;
        int count = (int)shapes.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic shp = shapes[i];
            try
            {
                if (string.Equals((string)shp.Name, name, StringComparison.OrdinalIgnoreCase))
                {
                    return shp;
                }
            }
            catch
            {
                // ignore shapes without a readable Name
            }
        }

        return null;
    }

    /// <summary>First shape on a slide whose Id equals <paramref name="id"/>, or null.</summary>
    public dynamic FindById(int slide, int id)
    {
        dynamic shapes = Prs.Slides[slide].Shapes;
        int count = (int)shapes.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic shp = shapes[i];
            try
            {
                if ((int)shp.Id == id)
                {
                    return shp;
                }
            }
            catch
            {
                // ignore shapes without a readable Id
            }
        }

        return null;
    }

    // ---------- text & font ----------

    /// <summary>Replace a shape's text.</summary>
    public void SetText(dynamic shp, string text) => shp.TextFrame.TextRange.Text = text;

    /// <summary>Read a shape's text (empty string if none).</summary>
    public string GetText(dynamic shp) => HasText(shp) ? (string)shp.TextFrame.TextRange.Text : string.Empty;

    /// <summary>
    /// Apply font changes. Pass only the parameters you want to change.
    /// <paramref name="colorBgr"/> is BGR (red = 0x0000FF).
    /// </summary>
    public void SetFont(
        dynamic shp,
        double? size = null,
        bool? bold = null,
        bool? italic = null,
        int? colorBgr = null,
        string name = null)
    {
        dynamic font = shp.TextFrame.TextRange.Font;
        if (size.HasValue)
        {
            font.Size = size.Value;
        }

        if (bold.HasValue)
        {
            font.Bold = bold.Value ? MsoTrue : MsoFalse;
        }

        if (italic.HasValue)
        {
            font.Italic = italic.Value ? MsoTrue : MsoFalse;
        }

        if (colorBgr.HasValue)
        {
            font.Color.RGB = colorBgr.Value;
        }

        if (name != null)
        {
            font.Name = name;
        }
    }

    // ---------- geometry ----------

    /// <summary>Move a shape (points).</summary>
    public void Move(dynamic shp, double left, double top)
    {
        shp.Left = left;
        shp.Top = top;
    }

    /// <summary>Resize a shape (points).</summary>
    public void Resize(dynamic shp, double width, double height)
    {
        shp.Width = width;
        shp.Height = height;
    }

    // ---------- appearance ----------

    /// <summary>Solid-fill a shape with a BGR color.</summary>
    public void SetFill(dynamic shp, int colorBgr)
    {
        shp.Fill.Solid();
        shp.Fill.ForeColor.RGB = colorBgr;
    }

    /// <summary>Set a shape's border color (BGR) and optional weight.</summary>
    public void SetBorder(dynamic shp, int colorBgr, double? weight = null)
    {
        shp.Line.ForeColor.RGB = colorBgr;
        if (weight.HasValue)
        {
            shp.Line.Weight = weight.Value;
        }
    }

    /// <summary>Set a slide's background fill to a BGR color.</summary>
    public void SetSlideBackground(int slide, int colorBgr)
    {
        dynamic bg = Prs.Slides[slide].Background;
        bg.Fill.Solid();
        bg.Fill.ForeColor.RGB = colorBgr;
    }

    // ---------- create ----------

    /// <summary>Add a textbox and return the new shape (points).</summary>
    public dynamic AddTextbox(int slide, string text, double left, double top, double width, double height)
    {
        dynamic shp = Prs.Slides[slide].Shapes.AddTextbox(
            MsoTextOrientationHorizontal, left, top, width, height);
        shp.TextFrame.TextRange.Text = text;
        return shp;
    }

    // ---------- notes ----------

    /// <summary>Set the speaker notes for a slide.</summary>
    public void SetNotes(int slide, string text)
    {
        dynamic notesPage = Prs.Slides[slide].NotesPage;
        int count = (int)notesPage.Shapes.Count;
        for (int i = 1; i <= count; i++)
        {
            dynamic shp = notesPage.Shapes[i];
            try
            {
                if ((int)shp.PlaceholderFormat.Type == 2)
                {
                    shp.TextFrame.TextRange.Text = text;
                    return;
                }
            }
            catch
            {
                // not a placeholder
            }
        }
    }

    private static bool HasText(dynamic shp)
    {
        try
        {
            return (int)shp.HasTextFrame == MsoTrue && (int)shp.TextFrame.HasText == MsoTrue;
        }
        catch
        {
            return false;
        }
    }
}

/// <summary>
/// Globals object for the <c>execute_template</c> (Level 3) path.
///
/// Unlike <c>execute_code</c> — where every request ships a fresh script string
/// that Roslyn must recompile (~200ms each) — the template path compiles ONE
/// parameterized script per action type, caches the compiled
/// <c>ScriptRunner&lt;object&gt;</c>, and re-invokes it with different values via
/// this globals instance. Values arrive through <see cref="A"/> (params, already
/// JSON-decoded) and <see cref="Shp"/> (the target shape resolved by the host),
/// so the script TEXT never changes → the compile cost is paid once per type.
/// </summary>
public sealed class TemplateGlobals
{
    /// <summary>The wrapped COM helper surface (same methods as CodeAct).</summary>
    public PptApi Api { get; }

    /// <summary>Target shape resolved by the host for target-based actions (may be null).</summary>
    public dynamic Shp;

    /// <summary>Per-call argument bag (action params plus <c>slide</c>).</summary>
    public Dictionary<string, object> A;

    public TemplateGlobals(PptApi api)
    {
        Api = api;
    }

    /// <summary>True when key <paramref name="k"/> is present and non-null.</summary>
    public bool Has(string k) => A != null && A.TryGetValue(k, out object v) && v != null;

    /// <summary>Read an int arg (tolerant of JSON long/double).</summary>
    public int I(string k) => Convert.ToInt32(A[k]);

    /// <summary>Read a double arg (tolerant of JSON long/int).</summary>
    public double D(string k) => Convert.ToDouble(A[k]);

    /// <summary>Read a bool arg.</summary>
    public bool B(string k) => Convert.ToBoolean(A[k]);

    /// <summary>Read a string arg (empty string if missing/null).</summary>
    public string S(string k) => A != null && A.TryGetValue(k, out object v) && v != null ? v.ToString() : string.Empty;
}
