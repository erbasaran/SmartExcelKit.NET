using System;
using System.Collections;
using System.Collections.Generic;
using SmartExcelKit.Core;

namespace SmartExcelKit;

/// <summary>
/// Represents a row in an Excel worksheet, providing access to row-level properties, cells, styles, and height.
/// </summary>
public sealed class ExcelRow : IEnumerable<ExcelCell>, IEquatable<ExcelRow>
{
    private readonly ExcelWorksheet _worksheet;

    /// <summary>
    /// Gets the 1-based row index.
    /// </summary>
    public int Index { get; }

    /// <summary>
    /// Gets the 1-based row number (alias for <see cref="Index"/>).
    /// </summary>
    public int RowNumber => Index;

    /// <summary>
    /// Gets the parent worksheet containing this row.
    /// </summary>
    public ExcelWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Gets or sets the height of this row in points.
    /// </summary>
    public double Height
    {
        get => _worksheet.GetRowHeight(Index);
        set => _worksheet.SetRowHeight(Index, value);
    }

    /// <summary>
    /// Gets or sets a value indicating whether this row is hidden.
    /// </summary>
    public bool Hidden
    {
        get => _worksheet.IsRowHidden(Index);
        set => _worksheet.SetRowHidden(Index, value);
    }

    /// <summary>
    /// Gets a value indicating whether this row contains any non-empty cell values, formulas, comments, or custom styles.
    /// </summary>
    public bool IsEmpty
    {
        get
        {
            int maxCol = Math.Max(1, _worksheet.MaxColumn);
            for (int col = 1; col <= maxCol; col++)
            {
                if (_worksheet.RawCells.TryGetValue(new CellAddress(Index, col), out var data) && data != null && !data.IsEmpty)
                {
                    return false;
                }
            }
            return true;
        }
    }

    /// <summary>
    /// Gets a lazy sequence of all cells in this row up to the maximum column in the worksheet.
    /// </summary>
    public IEnumerable<ExcelCell> Cells
    {
        get
        {
            int maxCol = Math.Max(1, _worksheet.MaxColumn);
            for (int col = 1; col <= maxCol; col++)
            {
                yield return _worksheet.Cell(Index, col);
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRow"/> class.
    /// </summary>
    /// <param name="worksheet">The parent worksheet.</param>
    /// <param name="index">The 1-based row index.</param>
    /// <exception cref="ArgumentNullException">Thrown if worksheet is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if row index is less than 1.</exception>
    public ExcelRow(ExcelWorksheet worksheet, int index)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        if (index < 1)
            throw new ArgumentOutOfRangeException(nameof(index), "Row index must be greater than or equal to 1.");
        Index = index;
    }

    /// <summary>
    /// Gets a cell in this row by 1-based column index.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <returns>An <see cref="ExcelCell"/> instance.</returns>
    public ExcelCell Cell(int column) => _worksheet.Cell(Index, column);

    /// <summary>
    /// Gets a cell in this row by 1-based column index.
    /// </summary>
    /// <param name="column">The 1-based column index.</param>
    /// <returns>An <see cref="ExcelCell"/> instance.</returns>
    public ExcelCell this[int column] => Cell(column);

    /// <summary>
    /// Returns a lazy enumeration of all used (non-empty) cells in this row.
    /// </summary>
    /// <returns>A sequence of used <see cref="ExcelCell"/> objects.</returns>
    public IEnumerable<ExcelCell> CellsUsed()
    {
        int maxCol = Math.Max(1, _worksheet.MaxColumn);
        for (int col = 1; col <= maxCol; col++)
        {
            var addr = new CellAddress(Index, col);
            if (_worksheet.RawCells.TryGetValue(addr, out var data) && data != null && !data.IsEmpty)
            {
                yield return _worksheet.Cell(addr);
            }
        }
    }

    /// <summary>
    /// Returns an array of raw cell values in this row up to <paramref name="maxColumns"/> (or worksheet MaxColumn if 0).
    /// </summary>
    /// <param name="maxColumns">Optional maximum 1-based column index (0 for worksheet MaxColumn).</param>
    public object?[] Values(int maxColumns = 0)
    {
        int limit = maxColumns > 0 ? maxColumns : Math.Max(1, _worksheet.MaxColumn);
        var result = new object?[limit];
        for (int col = 1; col <= limit; col++)
        {
            result[col - 1] = _worksheet.GetCellValue(new CellAddress(Index, col));
        }
        return result;
    }

    /// <summary>
    /// Returns an array of cell values in this row converted to <typeparamref name="T"/> up to <paramref name="maxColumns"/> (or worksheet MaxColumn if 0).
    /// </summary>
    /// <typeparam name="T">The target type.</typeparam>
    /// <param name="maxColumns">Optional maximum 1-based column index (0 for worksheet MaxColumn).</param>
    public T?[] Values<T>(int maxColumns = 0)
    {
        int limit = maxColumns > 0 ? maxColumns : Math.Max(1, _worksheet.MaxColumn);
        var result = new T?[limit];
        for (int col = 1; col <= limit; col++)
        {
            result[col - 1] = _worksheet.Cell(Index, col).GetValue<T>();
        }
        return result;
    }

    /// <summary>
    /// Alias for <see cref="Values(int)"/> returning an array of raw cell values.
    /// </summary>
    /// <param name="maxColumns">Optional maximum 1-based column index (0 for worksheet MaxColumn).</param>
    public object?[] ToArray(int maxColumns = 0) => Values(maxColumns);

    /// <summary>
    /// Returns a list of raw cell values in this row.
    /// </summary>
    /// <param name="maxColumns">Optional maximum 1-based column index (0 for worksheet MaxColumn).</param>
    public List<object?> ToList(int maxColumns = 0) => new(Values(maxColumns));

    /// <summary>
    /// Converts this row to a dictionary mapping header column names to cell values.
    /// Header names are retrieved from <paramref name="headerRow"/> (default 1). If no header row exists, column letters ("A", "B", ...) are used.
    /// </summary>
    /// <param name="headerRow">The 1-based row index containing column headers.</param>
    public Dictionary<string, object?> ToDictionary(int headerRow = 1)
    {
        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        int maxCol = Math.Max(1, _worksheet.MaxColumn);

        for (int col = 1; col <= maxCol; col++)
        {
            string header = string.Empty;
            if (headerRow >= 1 && headerRow != Index)
            {
                header = _worksheet.Cell(headerRow, col).GetString().Trim();
            }

            if (string.IsNullOrWhiteSpace(header))
            {
                header = CellAddress.GetColumnName(col);
            }

            // Ensure unique key in dictionary
            string uniqueHeader = header;
            int counter = 1;
            while (dict.ContainsKey(uniqueHeader))
            {
                uniqueHeader = $"{header}_{counter++}";
            }

            dict[uniqueHeader] = _worksheet.GetCellValue(new CellAddress(Index, col));
        }

        return dict;
    }

    /// <summary>
    /// Formats the cell values of this row as a single CSV formatted line.
    /// </summary>
    /// <param name="delimiter">The column delimiter character (default ',').</param>
    /// <param name="quoteCharacter">The quote character for escaping (default '"').</param>
    /// <param name="encoding">Optional encoding (unused for string return).</param>
    /// <param name="culture">Optional culture for formatting primitives.</param>
    public string ToCsv(char delimiter = ',', char quoteCharacter = '"', System.Text.Encoding? encoding = null, System.Globalization.CultureInfo? culture = null)
    {
        return Csv.CsvHelper.FormatRow(Values(), delimiter, quoteCharacter, culture);
    }

    /// <summary>
    /// Serializes this row to a JSON string using column headers.
    /// </summary>
    /// <param name="headerRow">The 1-based row index containing headers (default 1).</param>
    public string ToJson(int headerRow = 1)
    {
        var dict = ToDictionary(headerRow);
        return System.Text.Json.JsonSerializer.Serialize(dict);
    }

    /// <summary>
    /// Maps the cells of this row to a POCO instance of type <typeparamref name="T"/> using column header names from <paramref name="headerRow"/>.
    /// </summary>
    /// <typeparam name="T">The target POCO class type.</typeparam>
    /// <param name="headerRow">The 1-based header row index (default 1).</param>
    public T ToObject<T>(int headerRow = 1) where T : class, new()
    {
        var item = new T();
        int maxCol = Math.Max(1, _worksheet.MaxColumn);
        var props = Core.PropertyCache<T>.WriteableProperties;

        for (int col = 1; col <= maxCol; col++)
        {
            string header = _worksheet.Cell(headerRow, col).GetString().Trim().ToUpperInvariant();
            if (string.IsNullOrEmpty(header))
            {
                header = CellAddress.GetColumnName(col).ToUpperInvariant();
            }

            if (props.TryGetValue(header, out var prop))
            {
                var cellVal = _worksheet.GetCellValue(new CellAddress(Index, col));
                if (cellVal != null)
                {
                    var targetType = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
                    try
                    {
                        var converted = Convert.ChangeType(cellVal, targetType, System.Globalization.CultureInfo.InvariantCulture);
                        prop.SetValue(item, converted);
                    }
                    catch
                    {
                        if (targetType == typeof(string)) prop.SetValue(item, cellVal.ToString());
                    }
                }
            }
        }

        return item;
    }

    /// <summary>
    /// Returns true if this row contains no non-empty cell values, formulas, comments, or custom styles.
    /// </summary>
    public bool IsBlank() => IsEmpty;

    /// <summary>
    /// Returns the first non-empty cell in this row, or null if row is blank.
    /// </summary>
    public ExcelCell? FirstCellUsed()
    {
        int maxCol = Math.Max(1, _worksheet.MaxColumn);
        for (int col = 1; col <= maxCol; col++)
        {
            var addr = new CellAddress(Index, col);
            if (_worksheet.RawCells.TryGetValue(addr, out var data) && data != null && !data.IsEmpty)
            {
                return _worksheet.Cell(addr);
            }
        }
        return null;
    }

    /// <summary>
    /// Returns the last non-empty cell in this row, or null if row is blank.
    /// </summary>
    public ExcelCell? LastCellUsed()
    {
        int maxCol = Math.Max(1, _worksheet.MaxColumn);
        for (int col = maxCol; col >= 1; col--)
        {
            var addr = new CellAddress(Index, col);
            if (_worksheet.RawCells.TryGetValue(addr, out var data) && data != null && !data.IsEmpty)
            {
                return _worksheet.Cell(addr);
            }
        }
        return null;
    }

    /// <summary>
    /// Fast zero-allocation struct enumerator for iterating cells in an <see cref="ExcelRow"/>.
    /// </summary>
    public struct Enumerator : IEnumerator<ExcelCell>
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _row;
        private readonly int _maxCol;
        private int _currentCol;

        internal Enumerator(ExcelWorksheet worksheet, int row)
        {
            _worksheet = worksheet;
            _row = row;
            _maxCol = Math.Max(1, worksheet.MaxColumn);
            _currentCol = 0;
        }

        /// <inheritdoc />
        public bool MoveNext()
        {
            _currentCol++;
            return _currentCol <= _maxCol;
        }

        /// <inheritdoc />
        public ExcelCell Current => _worksheet.Cell(_row, _currentCol);

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public void Reset() => _currentCol = 0;

        /// <inheritdoc />
        public void Dispose() { }
    }

    /// <summary>
    /// Returns a zero-allocation struct enumerator that iterates through all cells in this row.
    /// </summary>
    public Enumerator GetEnumerator() => new(_worksheet, Index);

    IEnumerator<ExcelCell> IEnumerable<ExcelCell>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <inheritdoc />
    public bool Equals(ExcelRow? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return Index == other.Index && ReferenceEquals(_worksheet, other._worksheet);
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelRow other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (Index * 397) ^ _worksheet.GetHashCode();

    /// <inheritdoc />
    public override string ToString() => $"Row {Index}";
}
