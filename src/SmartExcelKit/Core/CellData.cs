namespace SmartExcelKit.Core;

/// <summary>
/// Internal data structure storing cell-level attributes. Stored inside worksheets to avoid heavy wrapper allocations.
/// </summary>
internal sealed class CellData
{
    /// <summary>
    /// Gets or sets the raw cell value (string, double, bool, DateTime, or null).
    /// </summary>
    public object? Value { get; set; }

    /// <summary>
    /// Gets or sets the cell formula (without the leading '=' sign).
    /// </summary>
    public string? Formula { get; set; }

    /// <summary>
    /// Gets or sets the style index reference from the workbook's StyleRegistry.
    /// </summary>
    public uint StyleId { get; set; }

    /// <summary>
    /// Gets or sets cell comments.
    /// </summary>
    public string? Comment { get; set; }

    /// <summary>
    /// Gets or sets the cell's hyperlink URL or reference.
    /// </summary>
    public string? Hyperlink { get; set; }

    /// <summary>
    /// Gets whether this cell data is empty (no value, formula, comment, hyperlink, or non-default style).
    /// </summary>
    public bool IsEmpty => Value == null && string.IsNullOrEmpty(Formula) && string.IsNullOrEmpty(Comment) && string.IsNullOrEmpty(Hyperlink) && StyleId == 0;
}
