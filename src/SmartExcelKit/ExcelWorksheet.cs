using SmartExcelKit.Core;
using SmartExcelKit.Styles;
using System.Data;
using System.Reflection;

namespace SmartExcelKit;

/// <summary>
/// Represents an Excel worksheet containing cells, rows, columns, styles, and settings.
/// </summary>
public sealed class ExcelWorksheet
{
    private readonly ExcelWorkbook _workbook;
    private string _name;
    private readonly Dictionary<CellAddress, CellData> _cells = [];
    private readonly Dictionary<int, double> _rowHeights = [];
    private readonly Dictionary<int, double> _columnWidths = [];
    private readonly HashSet<int> _hiddenRows = [];
    private readonly HashSet<int> _hiddenColumns = [];
    private readonly List<ExcelRangeAddress> _mergedRanges = [];

    /// <summary>
    /// Gets the workbook associated with this worksheet.
    /// </summary>
    public ExcelWorkbook Workbook => _workbook;

    /// <summary>
    /// Gets or sets the name of the worksheet.
    /// </summary>
    public string Name
    {
        get => _name;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(value));
            if (value.Length > 31)
                throw new ArgumentException("Worksheet name cannot exceed 31 characters.", nameof(value));
            _name = value;
        }
    }

    /// <summary>
    /// Gets or sets whether the worksheet shows gridlines.
    /// </summary>
    public bool ShowGridlines { get; set; } = true;

    /// <summary>
    /// Gets or sets the worksheet zoom percentage (between 10 and 400).
    /// </summary>
    public int Zoom { get; set; } = 100;

    /// <summary>
    /// Gets or sets the tab color in HEX format (e.g., "FF0000").
    /// </summary>
    public string? TabColor { get; set; }

    /// <summary>
    /// Gets or sets the number of frozen rows at the top.
    /// </summary>
    public int FreezeRows { get; set; }

    /// <summary>
    /// Gets or sets the number of frozen columns on the left.
    /// </summary>
    public int FreezeColumns { get; set; }

    /// <summary>
    /// Gets the list of merged cell ranges.
    /// </summary>
    public IReadOnlyList<ExcelRangeAddress> MergedRanges => _mergedRanges;

    /// <summary>
    /// Gets the sheet protection password hash (HEX format). Returns null if sheet is not protected.
    /// </summary>
    public string? ProtectionPasswordHash { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the worksheet is protected.
    /// </summary>
    public bool IsProtected => ProtectionPasswordHash != null;

    /// <summary>
    /// Protects the worksheet with a password using the Excel XOR hashing algorithm.
    /// </summary>
    /// <param name="password">The password to protect the sheet with.</param>
    public void Protect(string password)
    {
        if (string.IsNullOrEmpty(password))
        {
            ProtectionPasswordHash = "0000";
            return;
        }

        int hash = 0;
        for (int i = password.Length - 1; i >= 0; i--)
        {
            char c = password[i];
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= c;
        }
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= password.Length;
        hash ^= 0xCE4B;

        ProtectionPasswordHash = hash.ToString("X4");
    }

    /// <summary>
    /// Unprotects the worksheet.
    /// </summary>
    public void Unprotect()
    {
        ProtectionPasswordHash = null;
    }

    /// <summary>
    /// Gets the dictionary of raw cell data (internal use).
    /// </summary>
    internal IReadOnlyDictionary<CellAddress, CellData> RawCells => _cells;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelWorksheet"/> class.
    /// </summary>
    internal ExcelWorksheet(ExcelWorkbook workbook, string name)
    {
        _workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
        _name = name;
    }

    /// <summary>
    /// Retrieves a cell by its address.
    /// </summary>
    public ExcelCell Cell(CellAddress address) => new(this, address);

    /// <summary>
    /// Retrieves a cell by its row and column index (1-based).
    /// </summary>
    public ExcelCell Cell(int row, int column) => new(this, new CellAddress(row, column));

    /// <summary>
    /// Retrieves a cell by its string coordinate (e.g. "A1").
    /// </summary>
    public ExcelCell Cell(string address) => Cell(CellAddress.Parse(address));

    /// <summary>
    /// Retrieves a range by start and end coordinates.
    /// </summary>
    public ExcelRange Range(int startRow, int startColumn, int endRow, int endColumn) =>
        new(this, new ExcelRangeAddress(startRow, startColumn, endRow, endColumn));

    /// <summary>
    /// Retrieves a range by address string (e.g., "A1:B3").
    /// </summary>
    public ExcelRange Range(string rangeAddress) => new(this, ExcelRangeAddress.Parse(rangeAddress));

    /// <summary>
    /// Gets the maximum row index currently populated.
    /// </summary>
    public int MaxRow => _cells.Count == 0 ? 0 : _cells.Keys.Max(c => c.Row);

    /// <summary>
    /// Gets the maximum column index currently populated.
    /// </summary>
    public int MaxColumn => _cells.Count == 0 ? 0 : _cells.Keys.Max(c => c.Column);

    #region Internal Cell Data Operations

    private CellData GetOrCreateCellData(CellAddress address)
    {
        if (!_cells.TryGetValue(address, out var data))
        {
            data = new CellData();
            _cells[address] = data;
        }
        return data;
    }

    internal object? GetCellValue(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Value : null;
    }

    internal void SetCellValue(CellAddress address, object? value)
    {
        if (value == null)
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Value = null;
                data.Formula = null; // Clear formula if value is null'd
            }
            return;
        }
        GetOrCreateCellData(address).Value = value;
    }

    internal string? GetCellFormula(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Formula : null;
    }

    internal void SetCellFormula(CellAddress address, string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Formula = null;
            }
            return;
        }
        GetOrCreateCellData(address).Formula = formula!.TrimStart('=');
    }

    internal ExcelStyle GetCellStyle(CellAddress address)
    {
        if (_cells.TryGetValue(address, out var data))
        {
            return _workbook.StyleRegistry.GetStyle(data.StyleId);
        }
        return default;
    }

    internal void SetCellStyle(CellAddress address, ExcelStyle style)
    {
        uint styleId = _workbook.StyleRegistry.Register(style);
        GetOrCreateCellData(address).StyleId = styleId;
    }

    internal string? GetCellComment(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Comment : null;
    }

    internal void SetCellComment(CellAddress address, string? comment)
    {
        if (string.IsNullOrWhiteSpace(comment))
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Comment = null;
            }
            return;
        }
        GetOrCreateCellData(address).Comment = comment;
    }

    internal string? GetCellHyperlink(CellAddress address)
    {
        return _cells.TryGetValue(address, out var data) ? data.Hyperlink : null;
    }

    internal void SetCellHyperlink(CellAddress address, string? hyperlink)
    {
        if (string.IsNullOrWhiteSpace(hyperlink))
        {
            if (_cells.TryGetValue(address, out var data))
            {
                data.Hyperlink = null;
            }
            return;
        }
        GetOrCreateCellData(address).Hyperlink = hyperlink;
    }

    #endregion

    #region Layout Settings

    /// <summary>
    /// Sets the width of a specific column.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="width">The width value.</param>
    public void SetColumnWidth(int column, double width)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        if (width < 0) throw new ArgumentOutOfRangeException(nameof(width));
        _columnWidths[column] = width;
    }

    /// <summary>
    /// Gets the width of a specific column.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <returns>The column width.</returns>
    public double GetColumnWidth(int column)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        return _columnWidths.TryGetValue(column, out double w) ? w : 8.43; // Standard Excel column width
    }

    /// <summary>
    /// Sets the height of a specific row.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="height">The height value.</param>
    public void SetRowHeight(int row, double height)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (height < 0) throw new ArgumentOutOfRangeException(nameof(height));
        _rowHeights[row] = height;
    }

    /// <summary>
    /// Gets the height of a specific row.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>The row height.</returns>
    public double GetRowHeight(int row)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        return _rowHeights.TryGetValue(row, out double h) ? h : 15.0; // Standard Excel row height
    }

    /// <summary>
    /// Sets the visibility of a column.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <param name="hidden">True to hide the column, false to show it.</param>
    public void SetColumnHidden(int column, bool hidden)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column));
        if (hidden) _hiddenColumns.Add(column);
        else _hiddenColumns.Remove(column);
    }

    /// <summary>
    /// Gets whether a column is hidden.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <returns>True if hidden, false otherwise.</returns>
    public bool IsColumnHidden(int column) => _hiddenColumns.Contains(column);

    /// <summary>
    /// Sets the visibility of a row.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <param name="hidden">True to hide the row, false to show it.</param>
    public void SetRowHidden(int row, bool hidden)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row));
        if (hidden) _hiddenRows.Add(row);
        else _hiddenRows.Remove(row);
    }

    /// <summary>
    /// Gets whether a row is hidden.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>True if hidden, false otherwise.</returns>
    public bool IsRowHidden(int row) => _hiddenRows.Contains(row);

    /// <summary>
    /// Merges the specified range of cells.
    /// </summary>
    /// <param name="range">The range address to merge.</param>
    public void MergeCells(ExcelRangeAddress range)
    {
        if (!_mergedRanges.Contains(range))
        {
            _mergedRanges.Add(range);
        }
    }

    /// <summary>
    /// Unmerges the specified range of cells.
    /// </summary>
    /// <param name="range">The range address to unmerge.</param>
    public void UnmergeCells(ExcelRangeAddress range)
    {
        _mergedRanges.Remove(range);
    }

    /// <summary>
    /// Automatically fits the width of a column based on cell text content lengths.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    public void AutoFitColumn(int column)
    {
        int maxLen = 0;
        foreach (var cell in _cells)
        {
            if (cell.Key.Column == column && cell.Value.Value != null)
            {
                int len = cell.Value.Value.ToString()?.Length ?? 0;
                if (len > maxLen) maxLen = len;
            }
        }
        double width = Math.Max(8.43, (maxLen + 2) * 1.1); // Dynamic measurement approximation
        SetColumnWidth(column, width);
    }

    #endregion

    #region Data Import & Export

    /// <summary>
    /// Imports a strongly typed collection of objects into the worksheet starting at the specified row and column.
    /// </summary>
    public void Import<T>(IEnumerable<T> collection, int startRow = 1, int startColumn = 1)
    {
        if (collection == null) throw new ArgumentNullException(nameof(collection));
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                  .Where(p => p.CanRead)
                                  .ToList();

        // Write Headers
        for (int i = 0; i < properties.Count; i++)
        {
            Cell(startRow, startColumn + i).Value = properties[i].Name;
            Cell(startRow, startColumn + i).Style = new ExcelStyle(font: new ExcelFont(bold: true));
        }

        // Write Rows
        int currentRow = startRow + 1;
        foreach (var item in collection)
        {
            if (item == null) continue;
            for (int i = 0; i < properties.Count; i++)
            {
                var val = properties[i].GetValue(item);
                Cell(currentRow, startColumn + i).Value = val;
            }
            currentRow++;
        }
    }

    /// <summary>
    /// Exports the rows of the worksheet back to a strongly typed collection of objects.
    /// </summary>
    public IEnumerable<T> Export<T>(int startRow = 1, int startColumn = 1) where T : class, new()
    {
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        var properties = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                                  .Where(p => p.CanWrite)
                                  .ToDictionary(p => p.Name.ToUpperInvariant(), p => p);

        int maxRow = MaxRow;
        if (maxRow < startRow) yield break;

        // Read headers
        var colMap = new Dictionary<int, PropertyInfo>();
        int maxCol = MaxColumn;
        for (int col = startColumn; col <= maxCol; col++)
        {
            string header = Cell(startRow, col).GetString().Trim().ToUpperInvariant();
            if (properties.TryGetValue(header, out var prop))
            {
                colMap[col] = prop;
            }
        }

        for (int r = startRow + 1; r <= maxRow; r++)
        {
            var item = new T();
            bool hasAnyData = false;
            foreach (var mapping in colMap)
            {
                var cellVal = Cell(r, mapping.Key).Value;
                if (cellVal != null)
                {
                    hasAnyData = true;
                    var targetType = Nullable.GetUnderlyingType(mapping.Value.PropertyType) ?? mapping.Value.PropertyType;
                    try
                    {
                        var converted = Convert.ChangeType(cellVal, targetType);
                        mapping.Value.SetValue(item, converted);
                    }
                    catch
                    {
                        // Fallback conversion logic
                        if (targetType == typeof(string))
                        {
                            mapping.Value.SetValue(item, cellVal.ToString());
                        }
                    }
                }
            }
            if (hasAnyData)
            {
                yield return item;
            }
        }
    }

    /// <summary>
    /// Imports a DataTable into the worksheet.
    /// </summary>
    public void Import(DataTable dataTable, int startRow = 1, int startColumn = 1, bool includeHeader = true)
    {
        if (dataTable == null) throw new ArgumentNullException(nameof(dataTable));
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        int currentRow = startRow;

        if (includeHeader)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                Cell(currentRow, startColumn + col).Value = dataTable.Columns[col].ColumnName;
                Cell(currentRow, startColumn + col).Style = new ExcelStyle(font: new ExcelFont(bold: true));
            }
            currentRow++;
        }

        foreach (DataRow row in dataTable.Rows)
        {
            for (int col = 0; col < dataTable.Columns.Count; col++)
            {
                Cell(currentRow, startColumn + col).Value = row[col] == DBNull.Value ? null : row[col];
            }
            currentRow++;
        }
    }

    /// <summary>
    /// Exports sheet contents back to a DataTable.
    /// </summary>
    public DataTable ExportToDataTable(int startRow = 1, int startColumn = 1, bool hasHeader = true)
    {
        if (startRow < 1 || startColumn < 1) throw new ArgumentOutOfRangeException("Start coordinates must be >= 1.");

        var dt = new DataTable();
        int maxRow = MaxRow;
        int maxCol = MaxColumn;

        if (maxRow < startRow || maxCol < startColumn)
            return dt;

        int dataStartRow = startRow;

        // Establish columns
        if (hasHeader)
        {
            for (int col = startColumn; col <= maxCol; col++)
            {
                string headerName = Cell(startRow, col).GetString();
                if (string.IsNullOrWhiteSpace(headerName))
                {
                    headerName = $"Column_{col}";
                }
                dt.Columns.Add(headerName, typeof(string));
            }
            dataStartRow++;
        }
        else
        {
            for (int col = startColumn; col <= maxCol; col++)
            {
                dt.Columns.Add($"Column_{col}", typeof(string));
            }
        }

        // Fill data rows
        for (int r = dataStartRow; r <= maxRow; r++)
        {
            var rowValues = new string?[dt.Columns.Count];
            bool hasData = false;
            for (int c = 0; c < dt.Columns.Count; c++)
            {
                var val = Cell(r, startColumn + c).GetString();
                rowValues[c] = val;
                if (!string.IsNullOrEmpty(val)) hasData = true;
            }
            if (hasData)
            {
                dt.Rows.Add(rowValues);
            }
        }

        return dt;
    }

    #endregion
}
