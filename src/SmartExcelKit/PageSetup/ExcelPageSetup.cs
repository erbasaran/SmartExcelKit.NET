namespace SmartExcelKit.PageSetup;

/// <summary>
/// Specifies worksheet print orientation.
/// </summary>
public enum PageOrientation
{
    /// <summary>
    /// Default portrait orientation.
    /// </summary>
    Portrait,

    /// <summary>
    /// Horizontal landscape orientation.
    /// </summary>
    Landscape
}

/// <summary>
/// Specifies paper sizes for printing.
/// </summary>
public enum PaperSize
{
    /// <summary>
    /// Letter paper (8.5 x 11 in).
    /// </summary>
    Letter = 1,

    /// <summary>
    /// Legal paper (8.5 x 14 in).
    /// </summary>
    Legal = 5,

    /// <summary>
    /// Executive paper (7.25 x 10.5 in).
    /// </summary>
    Executive = 7,

    /// <summary>
    /// A3 paper (297 x 420 mm).
    /// </summary>
    A3 = 8,

    /// <summary>
    /// A4 paper (210 x 297 mm).
    /// </summary>
    A4 = 9,

    /// <summary>
    /// A5 paper (148 x 210 mm).
    /// </summary>
    A5 = 11,

    /// <summary>
    /// B4 paper (250 x 353 mm).
    /// </summary>
    B4 = 12,

    /// <summary>
    /// B5 paper (176 x 250 mm).
    /// </summary>
    B5 = 13
}

/// <summary>
/// Represents worksheet page setup, print options, margins, headers, and footers.
/// </summary>
public sealed class ExcelPageSetup
{
    /// <summary>
    /// Gets the parent worksheet.
    /// </summary>
    public ExcelWorksheet Worksheet { get; }

    /// <summary>
    /// Gets or sets page print orientation.
    /// </summary>
    public PageOrientation Orientation { get; set; } = PageOrientation.Portrait;

    /// <summary>
    /// Gets or sets paper size.
    /// </summary>
    public PaperSize PaperSize { get; set; } = PaperSize.A4;

    /// <summary>
    /// Gets or sets whether gridlines are printed.
    /// </summary>
    public bool PrintGridlines { get; set; } = false;

    /// <summary>
    /// Gets or sets whether row and column headings (A, B, C / 1, 2, 3) are printed.
    /// </summary>
    public bool PrintHeadings { get; set; } = false;

    /// <summary>
    /// Gets or sets print scaling percentage (10 to 400).
    /// </summary>
    public int Scale { get; set; } = 100;

    /// <summary>
    /// Gets or sets target page width count for fit-to-page printing (0 for unconstrained).
    /// </summary>
    public int FitToWidth { get; set; } = 0;

    /// <summary>
    /// Gets or sets target page height count for fit-to-page printing (0 for unconstrained).
    /// </summary>
    public int FitToHeight { get; set; } = 0;

    /// <summary>
    /// Gets or sets top margin in inches.
    /// </summary>
    public double TopMargin { get; set; } = 0.75;

    /// <summary>
    /// Gets or sets bottom margin in inches.
    /// </summary>
    public double BottomMargin { get; set; } = 0.75;

    /// <summary>
    /// Gets or sets left margin in inches.
    /// </summary>
    public double LeftMargin { get; set; } = 0.7;

    /// <summary>
    /// Gets or sets right margin in inches.
    /// </summary>
    public double RightMargin { get; set; } = 0.7;

    /// <summary>
    /// Gets or sets header margin in inches.
    /// </summary>
    public double HeaderMargin { get; set; } = 0.3;

    /// <summary>
    /// Gets or sets footer margin in inches.
    /// </summary>
    public double FooterMargin { get; set; } = 0.3;

    /// <summary>
    /// Gets or sets left header text.
    /// </summary>
    public string? HeaderLeft { get; set; }

    /// <summary>
    /// Gets or sets center header text.
    /// </summary>
    public string? HeaderCenter { get; set; }

    /// <summary>
    /// Gets or sets right header text.
    /// </summary>
    public string? HeaderRight { get; set; }

    /// <summary>
    /// Gets or sets left footer text.
    /// </summary>
    public string? FooterLeft { get; set; }

    /// <summary>
    /// Gets or sets center footer text.
    /// </summary>
    public string? FooterCenter { get; set; }

    /// <summary>
    /// Gets or sets right footer text.
    /// </summary>
    public string? FooterRight { get; set; }

    /// <summary>
    /// Gets or sets the explicit range address to print (e.g. "A1:G50").
    /// </summary>
    public string? PrintArea { get; set; }

    /// <summary>
    /// Gets or sets rows to repeat at top of each printed page (e.g. "$1:$2").
    /// </summary>
    public string? RepeatRows { get; set; }

    /// <summary>
    /// Gets or sets columns to repeat at left of each printed page (e.g. "$A:$B").
    /// </summary>
    public string? RepeatColumns { get; set; }

    internal ExcelPageSetup(ExcelWorksheet worksheet)
    {
        Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }
}
