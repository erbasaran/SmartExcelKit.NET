using System.Collections;
using SmartExcelKit.Core;

namespace SmartExcelKit.Tables;

/// <summary>
/// Represents an Excel Table (ListObject) in a worksheet.
/// </summary>
public sealed class ExcelTable : IEnumerable<ExcelTableColumn>
{
    private string _name;
    private ExcelRangeAddress _range;
    private readonly List<ExcelTableColumn> _columns = [];

    /// <summary>
    /// Gets the worksheet containing this table.
    /// </summary>
    public ExcelWorksheet Worksheet { get; }

    /// <summary>
    /// Gets or sets the unique name of the table.
    /// </summary>
    public string Name
    {
        get => _name;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Table name cannot be null or empty.", nameof(value));
            _name = value;
        }
    }

    /// <summary>
    /// Gets or sets the range address covered by the table.
    /// </summary>
    public ExcelRangeAddress Range
    {
        get => _range;
        set => _range = value;
    }

    /// <summary>
    /// Gets or sets whether the table displays a header row.
    /// </summary>
    public bool ShowHeaderRow { get; set; } = true;

    /// <summary>
    /// Gets or sets whether the table displays a totals row.
    /// </summary>
    public bool ShowTotalsRow { get; set; } = false;

    /// <summary>
    /// Gets or sets whether the table displays AutoFilter drop-downs on header cells.
    /// </summary>
    public bool ShowAutoFilter { get; set; } = true;

    /// <summary>
    /// Gets or sets whether row stripes styling is applied.
    /// </summary>
    public bool ShowRowStripes { get; set; } = true;

    /// <summary>
    /// Gets or sets whether column stripes styling is applied.
    /// </summary>
    public bool ShowColumnStripes { get; set; } = false;

    /// <summary>
    /// Gets or sets the Excel Table style name (e.g. "TableStyleMedium2", "TableStyleLight1").
    /// </summary>
    public string StyleName { get; set; } = "TableStyleMedium2";

    /// <summary>
    /// Gets the list of columns in the table.
    /// </summary>
    public IReadOnlyList<ExcelTableColumn> Columns => _columns;

    /// <summary>
    /// Gets a table column by name.
    /// </summary>
    public ExcelTableColumn this[string columnName]
    {
        get
        {
            var col = _columns.Find(c => string.Equals(c.Name, columnName, StringComparison.OrdinalIgnoreCase));
            if (col == null)
                throw new KeyNotFoundException($"Table column '{columnName}' was not found in table '{Name}'.");
            return col;
        }
    }

    /// <summary>
    /// Gets a table column by 0-based column index.
    /// </summary>
    public ExcelTableColumn this[int index] => _columns[index];

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelTable"/> class.
    /// </summary>
    internal ExcelTable(ExcelWorksheet worksheet, string name, ExcelRangeAddress range)
    {
        Worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _name = name;
        _range = range;

        InitializeColumnsFromRange();
    }

    private void InitializeColumnsFromRange()
    {
        int colIndex = 1;
        for (int c = _range.StartColumn; c <= _range.EndColumn; c++)
        {
            string colName = Worksheet.Cell(_range.StartRow, c).GetString();
            if (string.IsNullOrWhiteSpace(colName))
            {
                colName = $"Column{colIndex}";
            }

            // Ensure column names are unique within table
            string uniqueName = colName;
            int counter = 1;
            while (_columns.Exists(existing => string.Equals(existing.Name, uniqueName, StringComparison.OrdinalIgnoreCase)))
            {
                uniqueName = $"{colName}{counter++}";
            }

            _columns.Add(new ExcelTableColumn(this, colIndex++, uniqueName));
        }
    }

    /// <summary>
    /// Automatically expands the table bounds to include a new row.
    /// </summary>
    public void ExpandToIncludeRow()
    {
        _range = new ExcelRangeAddress(_range.StartRow, _range.StartColumn, _range.EndRow + 1, _range.EndColumn);
    }

    /// <inheritdoc />
    public IEnumerator<ExcelTableColumn> GetEnumerator() => _columns.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
