using System.Collections;
using SmartExcelKit.Styles;

namespace SmartExcelKit.Core;

/// <summary>
/// Represents a formatted text run within a rich text cell or comment.
/// </summary>
public sealed class RichTextRun
{
    /// <summary>
    /// Gets or sets the text content of the run.
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// Gets or sets the font styling applied to this run.
    /// </summary>
    public ExcelFont Font { get; set; }

    /// <summary>
    /// Initializes a new instance of the <see cref="RichTextRun"/> class.
    /// </summary>
    public RichTextRun(string text, ExcelFont font = default)
    {
        Text = text;
        Font = font;
    }
}

/// <summary>
/// Represents rich formatted cell text containing multiple styled text runs.
/// </summary>
public sealed class RichText : IEnumerable<RichTextRun>
{
    private readonly List<RichTextRun> _runs = [];

    /// <summary>
    /// Gets the number of text runs.
    /// </summary>
    public int Count => _runs.Count;

    /// <summary>
    /// Gets the run at the specified index.
    /// </summary>
    public RichTextRun this[int index] => _runs[index];

    /// <summary>
    /// Gets the read-only list of text runs.
    /// </summary>
    public IReadOnlyList<RichTextRun> Runs => _runs;

    /// <summary>
    /// Appends a plain text run with default font.
    /// </summary>
    public RichTextRun AddText(string text) => Add(text);

    /// <summary>
    /// Appends a new text run with optional font styling.
    /// </summary>
    public RichTextRun Add(string text, ExcelFont font = default)
    {
        var run = new RichTextRun(text, font);
        _runs.Add(run);
        return run;
    }

    /// <summary>
    /// Appends a bold text run.
    /// </summary>
    public RichTextRun AddBold(string text, double fontSize = 11, string fontColorHex = "000000")
    {
        return Add(text, new ExcelFont(bold: true, size: fontSize, color: fontColorHex));
    }

    /// <summary>
    /// Appends an italic text run.
    /// </summary>
    public RichTextRun AddItalic(string text, double fontSize = 11, string fontColorHex = "000000")
    {
        return Add(text, new ExcelFont(italic: true, size: fontSize, color: fontColorHex));
    }

    /// <summary>
    /// Clears all text runs.
    /// </summary>
    public void Clear() => _runs.Clear();

    /// <summary>
    /// Returns the combined plain text of all runs.
    /// </summary>
    public override string ToString() => string.Concat(_runs.Select(r => r.Text));

    /// <inheritdoc />
    public IEnumerator<RichTextRun> GetEnumerator() => _runs.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
