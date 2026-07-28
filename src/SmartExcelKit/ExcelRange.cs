using System.Collections;
using SmartExcelKit.Core;
using SmartExcelKit.Styles;

namespace SmartExcelKit;

/// <summary>
/// Represents a range of cells, allowing bulk updates, styles, formatting, merging, copying, and lazy iteration.
/// </summary>
public sealed class ExcelRange : IEnumerable<ExcelCell>, IEquatable<ExcelRange>
{
    private readonly ExcelWorksheet _worksheet;
    private readonly ExcelRangeAddress _address;

    /// <summary>
    /// Gets the range bounding address coordinates.
    /// </summary>
    public ExcelRangeAddress Address => _address;

    /// <summary>
    /// Gets the worksheet containing this range.
    /// </summary>
    public ExcelWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Gets a lazy sequence of all cells contained within this range.
    /// </summary>
    public IEnumerable<ExcelCell> Cells => this;

    /// <summary>
    /// Gets a lazy sequence of rows spanned by this range.
    /// </summary>
    public IEnumerable<ExcelRow> Rows
    {
        get
        {
            for (int r = _address.StartRow; r <= _address.EndRow; r++)
            {
                yield return _worksheet.Row(r);
            }
        }
    }

    /// <summary>
    /// Gets a lazy sequence of columns spanned by this range.
    /// </summary>
    public IEnumerable<ExcelColumn> Columns
    {
        get
        {
            for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
            {
                yield return _worksheet.Column(c);
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> class.
    /// </summary>
    /// <param name="worksheet">The parent worksheet.</param>
    /// <param name="address">The bounding box address.</param>
    internal ExcelRange(ExcelWorksheet worksheet, ExcelRangeAddress address)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        _address = address;
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> class from a string reference.
    /// </summary>
    /// <param name="worksheet">The parent worksheet.</param>
    /// <param name="rangeAddress">The range reference string (e.g. "A1:B10" or "C5").</param>
    public ExcelRange(ExcelWorksheet worksheet, string rangeAddress)
        : this(worksheet, ExcelRangeAddress.Parse(rangeAddress))
    {
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRange"/> class from bounding coordinates.
    /// </summary>
    /// <param name="worksheet">The parent worksheet.</param>
    /// <param name="firstRow">The 1-based start row.</param>
    /// <param name="firstColumn">The 1-based start column.</param>
    /// <param name="lastRow">The 1-based end row.</param>
    /// <param name="lastColumn">The 1-based end column.</param>
    public ExcelRange(ExcelWorksheet worksheet, int firstRow, int firstColumn, int lastRow, int lastColumn)
        : this(worksheet, new ExcelRangeAddress(firstRow, firstColumn, lastRow, lastColumn))
    {
    }

    /// <summary>
    /// Gets or sets the value for all cells within this range.
    /// When getting value for a 1x1 range, returns the cell value directly; otherwise returns a 2D matrix of values.
    /// </summary>
    public object? Value
    {
        get
        {
            if (_address.RowCount == 1 && _address.ColumnCount == 1)
            {
                return _worksheet.GetCellValue(new CellAddress(_address.StartRow, _address.StartColumn));
            }
            return GetValues();
        }
        set
        {
            for (int r = _address.StartRow; r <= _address.EndRow; r++)
            {
                for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
                {
                    _worksheet.SetCellValue(new CellAddress(r, c), value);
                }
            }
        }
    }

    /// <summary>
    /// Gets a cell inside this range by 1-based relative row and column offsets.
    /// </summary>
    /// <param name="row">1-based relative row index within range.</param>
    /// <param name="column">1-based relative column index within range.</param>
    /// <returns>An <see cref="ExcelCell"/> instance.</returns>
    public ExcelCell Cell(int row, int column)
    {
        return _worksheet.Cell(_address.StartRow + row - 1, _address.StartColumn + column - 1);
    }

    /// <summary>
    /// Gets a cell inside this range by cell address string relative or absolute.
    /// </summary>
    /// <param name="address">Cell reference string (e.g. "A1").</param>
    /// <returns>An <see cref="ExcelCell"/> instance.</returns>
    public ExcelCell Cell(string address) => _worksheet.Cell(address);

    /// <summary>
    /// Gets a cell inside this range by 1-based relative row and column offsets.
    /// </summary>
    public ExcelCell this[int row, int column] => Cell(row, column);

    /// <summary>
    /// Gets a cell inside this range or worksheet by address string (e.g. "A1").
    /// </summary>
    public ExcelCell this[string address] => Cell(address);

    /// <summary>
    /// Sets the formula for all cells within this range.
    /// </summary>
    public string? Formula
    {
        set
        {
            for (int r = _address.StartRow; r <= _address.EndRow; r++)
            {
                for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
                {
                    _worksheet.SetCellFormula(new CellAddress(r, c), value);
                }
            }
        }
    }

    /// <summary>
    /// Sets the style for all cells within this range.
    /// </summary>
    public ExcelStyle Style
    {
        set
        {
            for (int r = _address.StartRow; r <= _address.EndRow; r++)
            {
                for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
                {
                    _worksheet.SetCellStyle(new CellAddress(r, c), value);
                }
            }
        }
    }

    /// <summary>
    /// Gets a 2D array of values contained in this range.
    /// </summary>
    /// <returns>A 2D matrix of cell values indexed by [rowOffset, colOffset].</returns>
    public object?[,] GetValues()
    {
        int rowCount = _address.RowCount;
        int colCount = _address.ColumnCount;
        var matrix = new object?[rowCount, colCount];

        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                matrix[r, c] = _worksheet.GetCellValue(new CellAddress(_address.StartRow + r, _address.StartColumn + c));
            }
        }

        return matrix;
    }

    /// <summary>
    /// Gets a 2D array of values contained in this range (alias for <see cref="GetValues"/>).
    /// </summary>
    public object?[,] Values() => GetValues();

    /// <summary>
    /// Gets a 2D array of cell values in this range converted to type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    public T?[,] Values<T>()
    {
        int rowCount = _address.RowCount;
        int colCount = _address.ColumnCount;
        var matrix = new T?[rowCount, colCount];

        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                var cell = _worksheet.Cell(_address.StartRow + r, _address.StartColumn + c);
                matrix[r, c] = cell.GetValue<T>();
            }
        }

        return matrix;
    }

    /// <summary>
    /// Returns a 1D flattened array of raw cell values in this range.
    /// </summary>
    public object?[] ToArray()
    {
        int rowCount = _address.RowCount;
        int colCount = _address.ColumnCount;
        var array = new object?[rowCount * colCount];
        int idx = 0;

        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                array[idx++] = _worksheet.GetCellValue(new CellAddress(_address.StartRow + r, _address.StartColumn + c));
            }
        }

        return array;
    }

    /// <summary>
    /// Returns a 1D flattened list of raw cell values in this range.
    /// </summary>
    public List<object?> ToList() => new(ToArray());

    /// <summary>
    /// Formats the cell values of this range as CSV formatted text with newlines for rows.
    /// </summary>
    /// <param name="delimiter">The column delimiter character (default ',').</param>
    /// <param name="quoteCharacter">The quote character for escaping (default '"').</param>
    /// <param name="encoding">Optional encoding (unused for string return).</param>
    /// <param name="culture">Optional culture for formatting primitives.</param>
    public string ToCsv(char delimiter = ',', char quoteCharacter = '"', System.Text.Encoding? encoding = null, System.Globalization.CultureInfo? culture = null)
    {
        var sb = new System.Text.StringBuilder();
        int rowCount = _address.RowCount;
        int colCount = _address.ColumnCount;

        for (int r = 0; r < rowCount; r++)
        {
            var rowVals = new object?[colCount];
            for (int c = 0; c < colCount; c++)
            {
                rowVals[c] = _worksheet.GetCellValue(new CellAddress(_address.StartRow + r, _address.StartColumn + c));
            }
            if (r > 0) sb.AppendLine();
            sb.Append(Csv.CsvHelper.FormatRow(rowVals, delimiter, quoteCharacter, culture));
        }

        return sb.ToString();
    }

    /// <summary>
    /// Converts this range to a list of header-mapped dictionaries for each row.
    /// If <paramref name="hasHeader"/> is true, top row of range is used for header keys; otherwise column letters ("A", "B", ...) are used.
    /// </summary>
    /// <param name="hasHeader">Whether top row of range contains column headers (default true).</param>
    public List<Dictionary<string, object?>> ToDictionaryList(bool hasHeader = true)
    {
        var list = new List<Dictionary<string, object?>>();
        int startRow = _address.StartRow;
        int endRow = _address.EndRow;
        int startCol = _address.StartColumn;
        int endCol = _address.EndColumn;

        var headers = new List<string>();
        int dataStartRow = startRow;

        if (hasHeader)
        {
            for (int c = startCol; c <= endCol; c++)
            {
                string header = _worksheet.Cell(startRow, c).GetString().Trim();
                if (string.IsNullOrEmpty(header)) header = CellAddress.GetColumnName(c);
                headers.Add(header);
            }
            dataStartRow++;
        }
        else
        {
            for (int c = startCol; c <= endCol; c++)
            {
                headers.Add(CellAddress.GetColumnName(c));
            }
        }

        // Deduplicate header keys ONCE outside row loop
        var uniqueHeaders = new List<string>(headers.Count);
        var seenKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var h in headers)
        {
            string key = h;
            int counter = 1;
            while (seenKeys.Contains(key))
            {
                key = $"{h}_{counter++}";
            }
            seenKeys.Add(key);
            uniqueHeaders.Add(key);
        }

        for (int r = dataStartRow; r <= endRow; r++)
        {
            var dict = new Dictionary<string, object?>(uniqueHeaders.Count, StringComparer.OrdinalIgnoreCase);
            bool hasData = false;
            for (int c = 0; c < uniqueHeaders.Count; c++)
            {
                var val = _worksheet.GetCellValue(new CellAddress(r, startCol + c));
                dict[uniqueHeaders[c]] = val;
                if (val != null) hasData = true;
            }
            if (hasData)
            {
                list.Add(dict);
            }
        }

        return list;
    }

    /// <summary>
    /// Serializes this range to a JSON string representation.
    /// </summary>
    /// <param name="hasHeader">Whether top row contains headers for JSON property names (default true).</param>
    public string ToJson(bool hasHeader = true)
    {
        var dictList = ToDictionaryList(hasHeader);
        return System.Text.Json.JsonSerializer.Serialize(dictList);
    }

    /// <summary>
    /// Exports rows in this range to a <see cref="System.Data.DataTable"/>.
    /// </summary>
    /// <param name="hasHeader">Whether top row contains column names (default true).</param>
    public System.Data.DataTable ToDataTable(bool hasHeader = true)
    {
        return _worksheet.ExportToDataTable(_address.StartRow, _address.StartColumn, hasHeader);
    }

    /// <summary>
    /// Maps rows in this range to POCO objects of type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target POCO class type.</typeparam>
    /// <param name="hasHeader">Whether top row contains property headers (default true).</param>
    public IEnumerable<T> ToObjects<T>(bool hasHeader = true) where T : class, new()
    {
        int dataStartRow = hasHeader ? _address.StartRow : _address.StartRow - 1;
        return _worksheet.Export<T>(Math.Max(1, dataStartRow), _address.StartColumn, _address.EndRow);
    }

    /// <summary>
    /// Bulk writes a 2D array of values starting at the top-left cell of this range.
    /// </summary>
    /// <param name="values">The 2D matrix of cell values.</param>
    public void SetValues(object?[,] values)
    {
        if (values == null) throw new ArgumentNullException(nameof(values));

        int rows = values.GetLength(0);
        int cols = values.GetLength(1);

        for (int r = 0; r < rows; r++)
        {
            for (int c = 0; c < cols; c++)
            {
                int targetRow = _address.StartRow + r;
                int targetCol = _address.StartColumn + c;
                _worksheet.SetCellValue(new CellAddress(targetRow, targetCol), values[r, c]);
            }
        }
    }

    /// <summary>
    /// Bulk writes an enumerable of row value sequences starting at the top-left cell of this range.
    /// </summary>
    /// <param name="values">The sequence of row values.</param>
    public void SetValues(IEnumerable<IEnumerable<object?>> values)
    {
        if (values == null) throw new ArgumentNullException(nameof(values));

        int r = 0;
        foreach (var rowSeq in values)
        {
            if (rowSeq == null) continue;
            int c = 0;
            foreach (var item in rowSeq)
            {
                int targetRow = _address.StartRow + r;
                int targetCol = _address.StartColumn + c;
                _worksheet.SetCellValue(new CellAddress(targetRow, targetCol), item);
                c++;
            }
            r++;
        }
    }

    /// <summary>
    /// Copies all values, formulas, comments, hyperlinks, and styles from this range to the specified target range.
    /// </summary>
    /// <param name="target">The destination range.</param>
    public void CopyTo(ExcelRange target)
    {
        if (target == null) throw new ArgumentNullException(nameof(target));

        int rowCount = Math.Min(_address.RowCount, target._address.RowCount);
        int colCount = Math.Min(_address.ColumnCount, target._address.ColumnCount);

        for (int r = 0; r < rowCount; r++)
        {
            for (int c = 0; c < colCount; c++)
            {
                var srcAddr = new CellAddress(_address.StartRow + r, _address.StartColumn + c);
                var dstAddr = new CellAddress(target._address.StartRow + r, target._address.StartColumn + c);

                target._worksheet.SetCellValue(dstAddr, _worksheet.GetCellValue(srcAddr));
                target._worksheet.SetCellFormula(dstAddr, _worksheet.GetCellFormula(srcAddr));
                target._worksheet.SetCellComment(dstAddr, _worksheet.GetCellComment(srcAddr));
                target._worksheet.SetCellHyperlink(dstAddr, _worksheet.GetCellHyperlink(srcAddr));
                target._worksheet.SetCellStyle(dstAddr, _worksheet.GetCellStyle(srcAddr));
            }
        }
    }

    /// <summary>
    /// Merges the cells in this range.
    /// </summary>
    public void Merge()
    {
        _worksheet.MergeCells(_address);
    }

    /// <summary>
    /// Unmerges the cells in this range.
    /// </summary>
    public void Unmerge()
    {
        _worksheet.UnmergeCells(_address);
    }

    /// <summary>
    /// Clears values, formulas, comments, hyperlinks, and styles from all cells in this range.
    /// </summary>
    public void Clear()
    {
        ClearContents();
        ClearStyles();
        for (int r = _address.StartRow; r <= _address.EndRow; r++)
        {
            for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
            {
                var addr = new CellAddress(r, c);
                _worksheet.SetCellComment(addr, null);
                _worksheet.SetCellHyperlink(addr, null);
            }
        }
    }

    /// <summary>
    /// Resets formatting/styles of all cells in the range.
    /// </summary>
    public void ClearFormats()
    {
        Style = default;
    }

    /// <summary>
    /// Resets formatting/styles of all cells in the range (alias for <see cref="ClearFormats"/>).
    /// </summary>
    public void ClearStyles() => ClearFormats();

    /// <summary>
    /// Clears values and formulas of all cells in the range.
    /// </summary>
    public void ClearContents()
    {
        for (int r = _address.StartRow; r <= _address.EndRow; r++)
        {
            for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
            {
                var addr = new CellAddress(r, c);
                _worksheet.SetCellValue(addr, null);
                _worksheet.SetCellFormula(addr, null);
            }
        }
    }

    /// <summary>
    /// Automatically fits column widths to content for columns inside this range.
    /// </summary>
    public void AutoFitColumns()
    {
        for (int c = _address.StartColumn; c <= _address.EndColumn; c++)
        {
            _worksheet.AutoFitColumn(c);
        }
    }

    /// <summary>
    /// Fast zero-allocation struct enumerator for iterating cells in an <see cref="ExcelRange"/>.
    /// </summary>
    public struct Enumerator : IEnumerator<ExcelCell>
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _startRow;
        private readonly int _endRow;
        private readonly int _startCol;
        private readonly int _endCol;
        private int _currentRow;
        private int _currentCol;

        internal Enumerator(ExcelWorksheet worksheet, ExcelRangeAddress address)
        {
            _worksheet = worksheet;
            _startRow = address.StartRow;
            _endRow = address.EndRow;
            _startCol = address.StartColumn;
            _endCol = address.EndColumn;
            _currentRow = address.StartRow;
            _currentCol = address.StartColumn - 1;
        }

        /// <inheritdoc />
        public bool MoveNext()
        {
            _currentCol++;
            if (_currentCol > _endCol)
            {
                _currentCol = _startCol;
                _currentRow++;
            }
            return _currentRow <= _endRow;
        }

        /// <inheritdoc />
        public ExcelCell Current => _worksheet.Cell(_currentRow, _currentCol);

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public void Reset()
        {
            _currentRow = _startRow;
            _currentCol = _startCol - 1;
        }

        /// <inheritdoc />
        public void Dispose() { }
    }

    /// <summary>
    /// Enumerates all cells sequentially, row-by-row, column-by-column.
    /// </summary>
    public Enumerator GetEnumerator() => new(_worksheet, _address);

    IEnumerator<ExcelCell> IEnumerable<ExcelCell>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <inheritdoc />
    public bool Equals(ExcelRange? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return _address.Equals(other._address) && ReferenceEquals(_worksheet, other._worksheet);
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelRange other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (_address.GetHashCode() * 397) ^ _worksheet.GetHashCode();

    /// <inheritdoc />
    public override string ToString() => $"{_address.Address}";
}
