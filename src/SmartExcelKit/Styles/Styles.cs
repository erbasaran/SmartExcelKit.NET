namespace SmartExcelKit.Styles;

/// <summary>
/// Defines the border style.
/// </summary>
public enum ExcelBorderStyle
{
    /// <summary>No border.</summary>
    None,
    /// <summary>Thin border.</summary>
    Thin,
    /// <summary>Medium border.</summary>
    Medium,
    /// <summary>Dashed border.</summary>
    Dashed,
    /// <summary>Dotted border.</summary>
    Dotted,
    /// <summary>Thick border.</summary>
    Thick,
    /// <summary>Double border.</summary>
    Double,
    /// <summary>Hair border.</summary>
    Hair,
    /// <summary>Medium dashed border.</summary>
    MediumDashed,
    /// <summary>Dash-dot border.</summary>
    DashDot,
    /// <summary>Medium dash-dot border.</summary>
    MediumDashDot,
    /// <summary>Dash-dot-dot border.</summary>
    DashDotDot,
    /// <summary>Medium dash-dot-dot border.</summary>
    MediumDashDotDot,
    /// <summary>Slanted dash-dot border.</summary>
    SlantedDashDot
}

/// <summary>
/// Defines fill pattern types.
/// </summary>
public enum ExcelFillPatternType
{
    /// <summary>No fill pattern.</summary>
    None,
    /// <summary>Solid fill pattern.</summary>
    Solid,
    /// <summary>Dark gray pattern.</summary>
    DarkGray,
    /// <summary>Medium gray pattern.</summary>
    MediumGray,
    /// <summary>Light gray pattern.</summary>
    LightGray,
    /// <summary>12.5% gray pattern.</summary>
    Gray125,
    /// <summary>6.25% gray pattern.</summary>
    Gray0625,
    /// <summary>Dark horizontal stripes.</summary>
    DarkHorizontal,
    /// <summary>Dark vertical stripes.</summary>
    DarkVertical,
    /// <summary>Dark downward diagonal stripes.</summary>
    DarkDown,
    /// <summary>Dark upward diagonal stripes.</summary>
    DarkUp,
    /// <summary>Dark grid pattern.</summary>
    DarkGrid,
    /// <summary>Dark trellis pattern.</summary>
    DarkTrellis,
    /// <summary>Light horizontal stripes.</summary>
    LightHorizontal,
    /// <summary>Light vertical stripes.</summary>
    LightVertical,
    /// <summary>Light downward diagonal stripes.</summary>
    LightDown,
    /// <summary>Light upward diagonal stripes.</summary>
    LightUp,
    /// <summary>Light grid pattern.</summary>
    LightGrid,
    /// <summary>Light trellis pattern.</summary>
    LightTrellis
}

/// <summary>
/// Defines horizontal alignment options.
/// </summary>
public enum ExcelHorizontalAlignment
{
    /// <summary>General alignment.</summary>
    General,
    /// <summary>Left alignment.</summary>
    Left,
    /// <summary>Center alignment.</summary>
    Center,
    /// <summary>Right alignment.</summary>
    Right,
    /// <summary>Fill alignment.</summary>
    Fill,
    /// <summary>Justify alignment.</summary>
    Justify,
    /// <summary>Center continuous alignment.</summary>
    CenterContinuous,
    /// <summary>Distributed alignment.</summary>
    Distributed
}

/// <summary>
/// Defines vertical alignment options.
/// </summary>
public enum ExcelVerticalAlignment
{
    /// <summary>Top alignment.</summary>
    Top,
    /// <summary>Center alignment.</summary>
    Center,
    /// <summary>Bottom alignment.</summary>
    Bottom,
    /// <summary>Justify alignment.</summary>
    Justify,
    /// <summary>Distributed alignment.</summary>
    Distributed
}

/// <summary>
/// Represents individual border line properties.
/// </summary>
public readonly struct ExcelBorderItem : IEquatable<ExcelBorderItem>
{
    /// <summary>Gets the border style.</summary>
    public ExcelBorderStyle Style { get; }

    /// <summary>Gets the HEX color value of the border.</summary>
    public string? Color { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelBorderItem"/> struct.
    /// </summary>
    /// <param name="style">The border style.</param>
    /// <param name="color">The HEX color value.</param>
    public ExcelBorderItem(ExcelBorderStyle style, string? color = null)
    {
        Style = style;
        Color = color;
    }

    /// <inheritdoc />
    public bool Equals(ExcelBorderItem other) => Style == other.Style && Color == other.Color;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelBorderItem other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (int)Style ^ (Color?.GetHashCode() ?? 0);
}

/// <summary>
/// Represents border properties for all cell sides.
/// </summary>
public readonly struct ExcelBorder : IEquatable<ExcelBorder>
{
    /// <summary>Gets the left border properties.</summary>
    public ExcelBorderItem Left { get; }

    /// <summary>Gets the right border properties.</summary>
    public ExcelBorderItem Right { get; }

    /// <summary>Gets the top border properties.</summary>
    public ExcelBorderItem Top { get; }

    /// <summary>Gets the bottom border properties.</summary>
    public ExcelBorderItem Bottom { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelBorder"/> struct.
    /// </summary>
    /// <param name="left">The left border item.</param>
    /// <param name="right">The right border item.</param>
    /// <param name="top">The top border item.</param>
    /// <param name="bottom">The bottom border item.</param>
    public ExcelBorder(ExcelBorderItem left, ExcelBorderItem right, ExcelBorderItem top, ExcelBorderItem bottom)
    {
        Left = left;
        Right = right;
        Top = top;
        Bottom = bottom;
    }

    /// <inheritdoc />
    public bool Equals(ExcelBorder other) =>
        Left.Equals(other.Left) && Right.Equals(other.Right) && Top.Equals(other.Top) && Bottom.Equals(other.Bottom);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelBorder other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + Left.GetHashCode();
            hash = hash * 23 + Right.GetHashCode();
            hash = hash * 23 + Top.GetHashCode();
            hash = hash * 23 + Bottom.GetHashCode();
            return hash;
        }
    }
}

/// <summary>
/// Represents font properties.
/// </summary>
public readonly struct ExcelFont : IEquatable<ExcelFont>
{
    /// <summary>Gets the name of the font family.</summary>
    public string Name { get; }

    /// <summary>Gets the size of the font in points.</summary>
    public double Size { get; }

    /// <summary>Gets a value indicating whether the font is bold.</summary>
    public bool Bold { get; }

    /// <summary>Gets a value indicating whether the font is italic.</summary>
    public bool Italic { get; }

    /// <summary>Gets a value indicating whether the font is underlined.</summary>
    public bool Underline { get; }

    /// <summary>Gets the HEX color value of the font.</summary>
    public string? Color { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelFont"/> struct.
    /// </summary>
    /// <param name="name">The font family name.</param>
    /// <param name="size">The size in points.</param>
    /// <param name="bold">Whether bold styling is active.</param>
    /// <param name="italic">Whether italic styling is active.</param>
    /// <param name="underline">Whether underline styling is active.</param>
    /// <param name="color">The HEX color value.</param>
    public ExcelFont(string name = "Calibri", double size = 11, bool bold = false, bool italic = false, bool underline = false, string? color = null)
    {
        Name = name ?? "Calibri";
        Size = size;
        Bold = bold;
        Italic = italic;
        Underline = underline;
        Color = color;
    }

    /// <inheritdoc />
    public bool Equals(ExcelFont other) =>
        (Name ?? "Calibri") == (other.Name ?? "Calibri") &&
        Size.Equals(other.Size) &&
        Bold == other.Bold &&
        Italic == other.Italic &&
        Underline == other.Underline &&
        Color == other.Color;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelFont other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + (Name ?? "Calibri").GetHashCode();
            hash = hash * 23 + Size.GetHashCode();
            hash = hash * 23 + (Bold ? 1 : 0);
            hash = hash * 23 + (Italic ? 1 : 0);
            hash = hash * 23 + (Underline ? 1 : 0);
            hash = hash * 23 + (Color?.GetHashCode() ?? 0);
            return hash;
        }
    }
}

/// <summary>
/// Represents cell fill pattern properties.
/// </summary>
public readonly struct ExcelFill : IEquatable<ExcelFill>
{
    /// <summary>Gets the fill pattern type.</summary>
    public ExcelFillPatternType PatternType { get; }

    /// <summary>Gets the background color in HEX format.</summary>
    public string? BackgroundColor { get; }

    /// <summary>Gets the foreground color in HEX format.</summary>
    public string? ForegroundColor { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelFill"/> struct.
    /// </summary>
    /// <param name="patternType">The pattern type.</param>
    /// <param name="backgroundColor">The background HEX color.</param>
    /// <param name="foregroundColor">The foreground HEX color.</param>
    public ExcelFill(ExcelFillPatternType patternType = ExcelFillPatternType.None, string? backgroundColor = null, string? foregroundColor = null)
    {
        PatternType = patternType;
        BackgroundColor = backgroundColor;
        ForegroundColor = foregroundColor;
    }

    /// <inheritdoc />
    public bool Equals(ExcelFill other) =>
        PatternType == other.PatternType &&
        BackgroundColor == other.BackgroundColor &&
        ForegroundColor == other.ForegroundColor;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelFill other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + (int)PatternType;
            hash = hash * 23 + (BackgroundColor?.GetHashCode() ?? 0);
            hash = hash * 23 + (ForegroundColor?.GetHashCode() ?? 0);
            return hash;
        }
    }
}

/// <summary>
/// Represents text alignment settings inside a cell.
/// </summary>
public readonly struct ExcelAlignment : IEquatable<ExcelAlignment>
{
    /// <summary>Gets the horizontal alignment option.</summary>
    public ExcelHorizontalAlignment Horizontal { get; }

    /// <summary>Gets the vertical alignment option.</summary>
    public ExcelVerticalAlignment Vertical { get; }

    /// <summary>Gets a value indicating whether text should wrap.</summary>
    public bool WrapText { get; }

    /// <summary>Gets the indentation level.</summary>
    public int Indent { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelAlignment"/> struct.
    /// </summary>
    /// <param name="horizontal">The horizontal alignment.</param>
    /// <param name="vertical">The vertical alignment.</param>
    /// <param name="wrapText">Whether to wrap text.</param>
    /// <param name="indent">The indentation level.</param>
    public ExcelAlignment(ExcelHorizontalAlignment horizontal = ExcelHorizontalAlignment.General, ExcelVerticalAlignment vertical = ExcelVerticalAlignment.Bottom, bool wrapText = false, int indent = 0)
    {
        Horizontal = horizontal;
        Vertical = vertical;
        WrapText = wrapText;
        Indent = indent;
    }

    /// <inheritdoc />
    public bool Equals(ExcelAlignment other) =>
        Horizontal == other.Horizontal && Vertical == other.Vertical && WrapText == other.WrapText && Indent == other.Indent;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelAlignment other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + (int)Horizontal;
            hash = hash * 23 + (int)Vertical;
            hash = hash * 23 + (WrapText ? 1 : 0);
            hash = hash * 23 + Indent;
            return hash;
        }
    }
}

/// <summary>
/// Represents cell format pattern.
/// </summary>
public readonly struct ExcelNumberFormat : IEquatable<ExcelNumberFormat>
{
    /// <summary>Gets the format code pattern.</summary>
    public string FormatCode { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelNumberFormat"/> struct.
    /// </summary>
    /// <param name="formatCode">The format code string.</param>
    public ExcelNumberFormat(string formatCode = "General")
    {
        FormatCode = formatCode ?? "General";
    }

    /// <inheritdoc />
    public bool Equals(ExcelNumberFormat other) => (FormatCode ?? "General") == (other.FormatCode ?? "General");

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelNumberFormat other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (FormatCode ?? "General").GetHashCode();
}

/// <summary>
/// Composite styling structure containing font, fill, border, alignment, and number formatting.
/// </summary>
public readonly struct ExcelStyle : IEquatable<ExcelStyle>
{
    /// <summary>Gets the font component.</summary>
    public ExcelFont Font { get; }

    /// <summary>Gets the fill component.</summary>
    public ExcelFill Fill { get; }

    /// <summary>Gets the border component.</summary>
    public ExcelBorder Border { get; }

    /// <summary>Gets the text alignment component.</summary>
    public ExcelAlignment Alignment { get; }

    /// <summary>Gets the number format component.</summary>
    public ExcelNumberFormat NumberFormat { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelStyle"/> struct.
    /// </summary>
    /// <param name="font">The font component.</param>
    /// <param name="fill">The fill component.</param>
    /// <param name="border">The border component.</param>
    /// <param name="alignment">The alignment component.</param>
    /// <param name="numberFormat">The number format component.</param>
    public ExcelStyle(ExcelFont font = default, ExcelFill fill = default, ExcelBorder border = default, ExcelAlignment alignment = default, ExcelNumberFormat numberFormat = default)
    {
        Font = font;
        Fill = fill;
        Border = border;
        Alignment = alignment;
        NumberFormat = numberFormat;
    }

    /// <inheritdoc />
    public bool Equals(ExcelStyle other) =>
        Font.Equals(other.Font) &&
        Fill.Equals(other.Fill) &&
        Border.Equals(other.Border) &&
        Alignment.Equals(other.Alignment) &&
        NumberFormat.Equals(other.NumberFormat);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelStyle other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + Font.GetHashCode();
            hash = hash * 23 + Fill.GetHashCode();
            hash = hash * 23 + Border.GetHashCode();
            hash = hash * 23 + Alignment.GetHashCode();
            hash = hash * 23 + NumberFormat.GetHashCode();
            return hash;
        }
    }
}
