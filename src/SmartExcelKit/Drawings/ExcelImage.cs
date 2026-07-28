namespace SmartExcelKit.Drawings;

/// <summary>
/// Specifies common image format types.
/// </summary>
public enum ImageFormatType
{
    /// <summary>
    /// PNG image.
    /// </summary>
    Png,

    /// <summary>
    /// JPEG image.
    /// </summary>
    Jpeg,

    /// <summary>
    /// GIF image.
    /// </summary>
    Gif,

    /// <summary>
    /// BMP image.
    /// </summary>
    Bmp,

    /// <summary>
    /// SVG image.
    /// </summary>
    Svg
}

/// <summary>
/// Represents an image embedded in a worksheet.
/// </summary>
public sealed class ExcelImage
{
    /// <summary>
    /// Gets the unique identifier of the image.
    /// </summary>
    public string Id { get; } = Guid.NewGuid().ToString("N");

    /// <summary>
    /// Gets or sets the display name or title of the image.
    /// </summary>
    public string Name { get; set; } = "Picture";

    /// <summary>
    /// Gets the raw binary bytes of the image file.
    /// </summary>
    public byte[] ImageBytes { get; }

    /// <summary>
    /// Gets the format type of the image.
    /// </summary>
    public ImageFormatType Format { get; }

    /// <summary>
    /// Gets or sets the top-left 1-based row anchor.
    /// </summary>
    public int TopRow { get; set; }

    /// <summary>
    /// Gets or sets the top-left 1-based column anchor.
    /// </summary>
    public int LeftColumn { get; set; }

    /// <summary>
    /// Gets or sets the image width in pixels.
    /// </summary>
    public int Width { get; set; } = 300;

    /// <summary>
    /// Gets or sets the image height in pixels.
    /// </summary>
    public int Height { get; set; } = 200;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelImage"/> class.
    /// </summary>
    public ExcelImage(byte[] imageBytes, ImageFormatType format, int topRow, int leftColumn, int width = 300, int height = 200)
    {
        ImageBytes = imageBytes ?? throw new ArgumentNullException(nameof(imageBytes));
        Format = format;
        TopRow = topRow;
        LeftColumn = leftColumn;
        Width = width;
        Height = height;
    }

    /// <summary>
    /// Creates an image from a file path.
    /// </summary>
    public static ExcelImage FromFile(string filePath, int topRow, int leftColumn, int width = 300, int height = 200)
    {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be null or empty.", nameof(filePath));
        byte[] bytes = File.ReadAllBytes(filePath);
        string ext = Path.GetExtension(filePath).ToUpperInvariant();
        ImageFormatType fmt = ext switch
        {
            ".PNG" => ImageFormatType.Png,
            ".JPG" or ".JPEG" => ImageFormatType.Jpeg,
            ".GIF" => ImageFormatType.Gif,
            ".BMP" => ImageFormatType.Bmp,
            ".SVG" => ImageFormatType.Svg,
            _ => ImageFormatType.Png
        };
        return new ExcelImage(bytes, fmt, topRow, leftColumn, width, height);
    }
}
