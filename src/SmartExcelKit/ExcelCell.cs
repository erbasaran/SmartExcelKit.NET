using SmartExcelKit.Core;
using SmartExcelKit.Styles;

namespace SmartExcelKit;

/// <summary>
/// Represents a cell wrapper providing clean, intuitive access to values, formulas, comments, and styles.
/// </summary>
public sealed class ExcelCell
{
    private readonly ExcelWorksheet _worksheet;
    private readonly CellAddress _address;

    /// <summary>
    /// Gets the cell address coordinates.
    /// </summary>
    public CellAddress Address => _address;

    /// <summary>
    /// Gets the worksheet containing this cell.
    /// </summary>
    public ExcelWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelCell"/> class.
    /// </summary>
    internal ExcelCell(ExcelWorksheet worksheet, CellAddress address)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _address = address;
    }

    /// <summary>
    /// Gets or sets the cell value.
    /// </summary>
    public object? Value
    {
        get => _worksheet.GetCellValue(_address);
        set => _worksheet.SetCellValue(_address, value);
    }

    /// <summary>
    /// Gets or sets the cell formula (without the leading '=' sign).
    /// </summary>
    public string? Formula
    {
        get => _worksheet.GetCellFormula(_address);
        set => _worksheet.SetCellFormula(_address, value);
    }

    /// <summary>
    /// Gets or sets the cell style.
    /// </summary>
    public ExcelStyle Style
    {
        get => _worksheet.GetCellStyle(_address);
        set => _worksheet.SetCellStyle(_address, value);
    }

    /// <summary>
    /// Gets or sets the comment associated with this cell.
    /// </summary>
    public string? Comment
    {
        get => _worksheet.GetCellComment(_address);
        set => _worksheet.SetCellComment(_address, value);
    }

    /// <summary>
    /// Gets or sets the hyperlink associated with this cell.
    /// </summary>
    public string? Hyperlink
    {
        get => _worksheet.GetCellHyperlink(_address);
        set => _worksheet.SetCellHyperlink(_address, value);
    }

    /// <summary>
    /// Returns the string representation of the value.
    /// </summary>
    public string GetString() => Value?.ToString() ?? string.Empty;

    /// <summary>
    /// Safely attempts to convert the cell value to a double.
    /// </summary>
    public double GetDouble()
    {
        var val = Value;
        if (val is double d) return d;
        if (val is float f) return f;
        if (val is int i) return i;
        if (val is decimal dec) return (double)dec;
        if (val is string s && double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsed)) return parsed;
        return 0.0;
    }

    /// <summary>
    /// Safely attempts to convert the cell value to a boolean.
    /// </summary>
    public bool GetBoolean()
    {
        var val = Value;
        if (val is bool b) return b;
        if (val is string s && bool.TryParse(s, out bool parsed)) return parsed;
        if (val is double d) return d != 0.0;
        if (val is int i) return i != 0;
        return false;
    }

    /// <summary>
    /// Safely attempts to convert the cell value to a DateTime.
    /// </summary>
    public DateTime GetDateTime()
    {
        var val = Value;
        if (val is DateTime dt) return dt;
        if (val is string s && DateTime.TryParse(s, out DateTime parsed)) return parsed;
        if (val is double d) return DateTime.FromOADate(d);
        return DateTime.MinValue;
    }
}
