namespace SmartExcelKit.Core;

/// <summary>
/// Represents a cell comment / note.
/// </summary>
public sealed class ExcelComment
{
    /// <summary>
    /// Gets or sets the author name of the comment.
    /// </summary>
    public string Author { get; set; } = "SmartExcelKit";

    /// <summary>
    /// Gets or sets the plain text comment content.
    /// </summary>
    public string Text { get; set; }

    /// <summary>
    /// Gets or sets optional rich text formatted content.
    /// </summary>
    public RichText? RichText { get; set; }

    /// <summary>
    /// Gets or sets whether the comment is permanently visible on the sheet.
    /// </summary>
    public bool IsVisible { get; set; } = false;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelComment"/> class.
    /// </summary>
    public ExcelComment(string text, string author = "SmartExcelKit")
    {
        Text = text;
        Author = author;
    }
}
