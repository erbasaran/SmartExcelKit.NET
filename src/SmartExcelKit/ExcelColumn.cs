using System.Collections;
using SmartExcelKit.Core;

namespace SmartExcelKit;

/// <summary>
/// Represents a column in an Excel worksheet, providing access to column-level properties, cells, styles, and width.
/// </summary>
public sealed class ExcelColumn : IEnumerable<ExcelCell>, IEquatable<ExcelColumn>
{
    private readonly ExcelWorksheet _worksheet;

    /// <summary>
    /// Gets the 1-based column index.
    /// </summary>
    public int Index { get; }

    /// <summary>
    /// Gets the 1-based column number (alias for <see cref="Index"/>).
    /// </summary>
    public int ColumnNumber => Index;

    /// <summary>
    /// Gets the alphabetical column letter (e.g., "A", "BC").
    /// </summary>
    public string Letter => CellAddress.GetColumnName(Index);

    /// <summary>
    /// Gets the parent worksheet containing this column.
    /// </summary>
    public ExcelWorksheet Worksheet => _worksheet;

    /// <summary>
    /// Gets or sets the width of this column.
    /// </summary>
    public double Width
    {
        get => _worksheet.GetColumnWidth(Index);
        set => _worksheet.SetColumnWidth(Index, value);
    }

    /// <summary>
    /// Gets or sets a value indicating whether this column is hidden.
    /// </summary>
    public bool Hidden
    {
        get => _worksheet.IsColumnHidden(Index);
        set => _worksheet.SetColumnHidden(Index, value);
    }

    /// <summary>
    /// Gets a value indicating whether this column contains any non-empty cell values, formulas, comments, or custom styles.
    /// </summary>
    public bool IsEmpty
    {
        get
        {
            int maxRow = Math.Max(1, _worksheet.MaxRow);
            for (int row = 1; row <= maxRow; row++)
            {
                if (_worksheet.RawCells.TryGetValue(new CellAddress(row, Index), out var data) && data != null && !data.IsEmpty)
                {
                    return false;
                }
            }
            return true;
        }
    }

    /// <summary>
    /// Gets a lazy sequence of all cells in this column up to the maximum row in the worksheet.
    /// </summary>
    public IEnumerable<ExcelCell> Cells
    {
        get
        {
            int maxRow = Math.Max(1, _worksheet.MaxRow);
            for (int row = 1; row <= maxRow; row++)
            {
                yield return _worksheet.Cell(row, Index);
            }
        }
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelColumn"/> class.
    /// </summary>
    /// <param name="worksheet">The parent worksheet.</param>
    /// <param name="index">The 1-based column index.</param>
    /// <exception cref="ArgumentNullException">Thrown if worksheet is null.</exception>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if column index is less than 1.</exception>
    public ExcelColumn(ExcelWorksheet worksheet, int index)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
        if (index < 1)
            throw new ArgumentOutOfRangeException(nameof(index), "Column index must be greater than or equal to 1.");
        Index = index;
    }

    /// <summary>
    /// Gets a cell in this column by 1-based row index.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>An <see cref="ExcelCell"/> instance.</returns>
    public ExcelCell Cell(int row) => _worksheet.Cell(row, Index);

    /// <summary>
    /// Gets a cell in this column by 1-based row index.
    /// </summary>
    /// <param name="row">The 1-based row index.</param>
    /// <returns>An <see cref="ExcelCell"/> instance.</returns>
    public ExcelCell this[int row] => Cell(row);

    /// <summary>
    /// Returns a lazy enumeration of all used (non-empty) cells in this column.
    /// </summary>
    /// <returns>A sequence of used <see cref="ExcelCell"/> objects.</returns>
    public IEnumerable<ExcelCell> CellsUsed()
    {
        int maxRow = Math.Max(1, _worksheet.MaxRow);
        for (int row = 1; row <= maxRow; row++)
        {
            var addr = new CellAddress(row, Index);
            if (_worksheet.RawCells.TryGetValue(addr, out var data) && data != null && !data.IsEmpty)
            {
                yield return _worksheet.Cell(addr);
            }
        }
    }

    /// <summary>
    /// Automatically fits this column width based on content.
    /// </summary>
    public void AutoFit() => _worksheet.AutoFitColumn(Index);

    /// <summary>
    /// Fast zero-allocation struct enumerator for iterating cells in an <see cref="ExcelColumn"/>.
    /// </summary>
    public struct Enumerator : IEnumerator<ExcelCell>
    {
        private readonly ExcelWorksheet _worksheet;
        private readonly int _column;
        private readonly int _maxRow;
        private int _currentRow;

        internal Enumerator(ExcelWorksheet worksheet, int column)
        {
            _worksheet = worksheet;
            _column = column;
            _maxRow = Math.Max(1, worksheet.MaxRow);
            _currentRow = 0;
        }

        /// <inheritdoc />
        public bool MoveNext()
        {
            _currentRow++;
            return _currentRow <= _maxRow;
        }

        /// <inheritdoc />
        public ExcelCell Current => _worksheet.Cell(_currentRow, _column);

        object IEnumerator.Current => Current;

        /// <inheritdoc />
        public void Reset() => _currentRow = 0;

        /// <inheritdoc />
        public void Dispose() { }
    }

    /// <summary>
    /// Returns a zero-allocation struct enumerator that iterates through all cells in this column.
    /// </summary>
    public Enumerator GetEnumerator() => new(_worksheet, Index);

    IEnumerator<ExcelCell> IEnumerable<ExcelCell>.GetEnumerator() => GetEnumerator();

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <inheritdoc />
    public bool Equals(ExcelColumn? other)
    {
        if (other is null) return false;
        if (ReferenceEquals(this, other)) return true;
        return Index == other.Index && ReferenceEquals(_worksheet, other._worksheet);
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelColumn other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() => (Index * 397) ^ _worksheet.GetHashCode();

    /// <inheritdoc />
    public override string ToString() => $"Column {Letter} ({Index})";
}
