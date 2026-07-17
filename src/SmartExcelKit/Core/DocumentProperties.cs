namespace SmartExcelKit.Core;

/// <summary>
/// Represents standard document metadata properties for the workbook.
/// </summary>
public sealed class DocumentProperties
{
    /// <summary>
    /// Gets or sets the title of the document.
    /// </summary>
    public string? Title { get; set; }

    /// <summary>
    /// Gets or sets the author/creator of the document.
    /// </summary>
    public string? Author { get; set; }

    /// <summary>
    /// Gets or sets the subject of the document.
    /// </summary>
    public string? Subject { get; set; }

    /// <summary>
    /// Gets or sets the keywords list associated with the document.
    /// </summary>
    public string? Keywords { get; set; }

    /// <summary>
    /// Gets or sets the description of the document contents.
    /// </summary>
    public string? Description { get; set; }

    /// <summary>
    /// Gets or sets the document category.
    /// </summary>
    public string? Category { get; set; }

    /// <summary>
    /// Gets or sets the company name.
    /// </summary>
    public string? Company { get; set; }

    /// <summary>
    /// Gets or sets the date and time when the document was created.
    /// </summary>
    public DateTime Created { get; set; } = DateTime.UtcNow;

    /// <summary>
    /// Gets or sets the date and time when the document was last modified.
    /// </summary>
    public DateTime Modified { get; set; } = DateTime.UtcNow;
}
