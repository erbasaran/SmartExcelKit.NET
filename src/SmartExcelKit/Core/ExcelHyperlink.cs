namespace SmartExcelKit.Core;

/// <summary>
/// Specifies the type of hyperlink.
/// </summary>
public enum HyperlinkType
{
    /// <summary>
    /// External web URL (http:// or https://).
    /// </summary>
    ExternalUrl,

    /// <summary>
    /// Internal workbook cell or sheet reference (e.g. "'Sheet2'!A1").
    /// </summary>
    InternalReference,

    /// <summary>
    /// Email mailto hyperlink.
    /// </summary>
    Email,

    /// <summary>
    /// Local or network file path.
    /// </summary>
    File
}

/// <summary>
/// Represents a hyperlink attached to a cell.
/// </summary>
public sealed class ExcelHyperlink
{
    /// <summary>
    /// Gets or sets the target URI or cell reference string.
    /// </summary>
    public string Target { get; set; }

    /// <summary>
    /// Gets or sets the hyperlink type.
    /// </summary>
    public HyperlinkType HyperlinkType { get; set; }

    /// <summary>
    /// Gets or sets optional mouse-over tooltip text.
    /// </summary>
    public string? Tooltip { get; set; }

    /// <summary>
    /// Gets or sets optional display text override.
    /// </summary>
    public string? DisplayText { get; set; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelHyperlink"/> class.
    /// </summary>
    public ExcelHyperlink(string target, HyperlinkType hyperlinkType = HyperlinkType.ExternalUrl, string? tooltip = null)
    {
        if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("Hyperlink target cannot be null or empty.", nameof(target));
        Target = target;
        HyperlinkType = hyperlinkType;
        Tooltip = tooltip;
    }

    /// <summary>
    /// Creates an external web URL hyperlink.
    /// </summary>
    public static ExcelHyperlink External(string url, string? tooltip = null)
    {
        return new ExcelHyperlink(url, HyperlinkType.ExternalUrl, tooltip);
    }

    /// <summary>
    /// Creates an internal cell reference hyperlink.
    /// </summary>
    public static ExcelHyperlink Internal(string sheetName, string cellAddress, string? tooltip = null)
    {
        return new ExcelHyperlink($"'{sheetName}'!{cellAddress}", HyperlinkType.InternalReference, tooltip);
    }

    /// <summary>
    /// Creates an email mailto hyperlink.
    /// </summary>
    public static ExcelHyperlink Email(string emailAddress, string? subject = null, string? tooltip = null)
    {
        string target = string.IsNullOrEmpty(subject) ? $"mailto:{emailAddress}" : $"mailto:{emailAddress}?subject={Uri.EscapeDataString(subject!)}";
        return new ExcelHyperlink(target, HyperlinkType.Email, tooltip);
    }
}
