namespace SmartExcelKit.Tables;

/// <summary>
/// Represents a column in an <see cref="ExcelTable"/>.
/// </summary>
public sealed class ExcelTableColumn
{
    private string _name;

    /// <summary>
    /// Gets the parent table.
    /// </summary>
    public ExcelTable Table { get; }

    /// <summary>
    /// Gets the 1-based index of this column within the table.
    /// </summary>
    public int Index { get; }

    /// <summary>
    /// Gets or sets the header name of the column.
    /// </summary>
    public string Name
    {
        get => _name;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Table column name cannot be null or empty.", nameof(value));
            _name = value;
        }
    }

    /// <summary>
    /// Gets or sets the totals row aggregation function.
    /// </summary>
    public TotalsRowFunction TotalsRowFunction { get; set; } = TotalsRowFunction.None;

    /// <summary>
    /// Gets or sets explicit text for the totals row cell when <see cref="TotalsRowFunction"/> is <see cref="TotalsRowFunction.None"/>.
    /// </summary>
    public string? TotalsRowLabel { get; set; }

    /// <summary>
    /// Gets or sets a custom formula for the totals row cell when <see cref="TotalsRowFunction"/> is <see cref="TotalsRowFunction.Custom"/>.
    /// </summary>
    public string? TotalsRowFormula { get; set; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelTableColumn"/> class.
    /// </summary>
    internal ExcelTableColumn(ExcelTable table, int index, string name)
    {
        Table = table ?? throw new ArgumentNullException(nameof(table));
        Index = index;
        _name = name;
    }
}
