namespace SmartExcelKit.Core;

/// <summary>
/// Represents a range of cells in a worksheet defined by bounding box coordinates.
/// </summary>
public readonly struct ExcelRangeAddress : IEquatable<ExcelRangeAddress>
{
    /// <summary>
    /// Gets the start row index (1-based).
    /// </summary>
    public int StartRow { get; }

    /// <summary>
    /// Gets the start column index (1-based).
    /// </summary>
    public int StartColumn { get; }

    /// <summary>
    /// Gets the end row index (1-based).
    /// </summary>
    public int EndRow { get; }

    /// <summary>
    /// Gets the end column index (1-based).
    /// </summary>
    public int EndColumn { get; }

    /// <summary>
    /// Gets the address in range notation (e.g. "A1:B3").
    /// </summary>
    public string Address
    {
        get
        {
            if (StartRow == EndRow && StartColumn == EndColumn)
            {
                return new CellAddress(StartRow, StartColumn).Address;
            }
            return $"{new CellAddress(StartRow, StartColumn).Address}:{new CellAddress(EndRow, EndColumn).Address}";
        }
    }

    /// <summary>
    /// Gets the total number of rows in the range.
    /// </summary>
    public int RowCount => EndRow - StartRow + 1;

    /// <summary>
    /// Gets the total number of columns in the range.
    /// </summary>
    public int ColumnCount => EndColumn - StartColumn + 1;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelRangeAddress"/> struct.
    /// </summary>
    /// <param name="startRow">The start row number (1-based).</param>
    /// <param name="startColumn">The start column number (1-based).</param>
    /// <param name="endRow">The end row number (1-based).</param>
    /// <param name="endColumn">The end column number (1-based).</param>
    public ExcelRangeAddress(int startRow, int startColumn, int endRow, int endColumn)
    {
        StartRow = Math.Min(startRow, endRow);
        EndRow = Math.Max(startRow, endRow);
        StartColumn = Math.Min(startColumn, endColumn);
        EndColumn = Math.Max(startColumn, endColumn);

        if (StartRow < 1 || StartColumn < 1)
            throw new ArgumentOutOfRangeException("Row and Column numbers must be greater than or equal to 1.");
    }

    /// <summary>
    /// Parses a range address string (e.g. "A1:C5", "B2").
    /// </summary>
    /// <param name="rangeAddress">The range address string.</param>
    /// <returns>A new <see cref="ExcelRangeAddress"/> instance.</returns>
    public static ExcelRangeAddress Parse(string rangeAddress)
    {
        if (string.IsNullOrWhiteSpace(rangeAddress))
            throw new ArgumentException("Range address cannot be null or empty.", nameof(rangeAddress));

        int colonIndex = rangeAddress.IndexOf(':');
        if (colonIndex < 0)
        {
            var cell = CellAddress.Parse(rangeAddress);
            return new ExcelRangeAddress(cell.Row, cell.Column, cell.Row, cell.Column);
        }

        var startCell = CellAddress.Parse(rangeAddress.Substring(0, colonIndex));
        var endCell = CellAddress.Parse(rangeAddress.Substring(colonIndex + 1));
        return new ExcelRangeAddress(startCell.Row, startCell.Column, endCell.Row, endCell.Column);
    }

    /// <summary>
    /// Returns whether the range contains the specified cell.
    /// </summary>
    public bool Contains(CellAddress cellAddress)
    {
        return cellAddress.Row >= StartRow && cellAddress.Row <= EndRow &&
               cellAddress.Column >= StartColumn && cellAddress.Column <= EndColumn;
    }

    /// <summary>
    /// Returns whether the range intersects with another range.
    /// </summary>
    public bool Intersects(ExcelRangeAddress other)
    {
        return !(EndRow < other.StartRow || StartRow > other.EndRow ||
                 EndColumn < other.StartColumn || StartColumn > other.EndColumn);
    }

    /// <inheritdoc />
    public bool Equals(ExcelRangeAddress other)
    {
        return StartRow == other.StartRow &&
               StartColumn == other.StartColumn &&
               EndRow == other.EndRow &&
               EndColumn == other.EndColumn;
    }

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is ExcelRangeAddress other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        unchecked
        {
            int hash = 17;
            hash = hash * 23 + StartRow;
            hash = hash * 23 + StartColumn;
            hash = hash * 23 + EndRow;
            hash = hash * 23 + EndColumn;
            return hash;
        }
    }

    /// <summary>Compares two ExcelRangeAddress values for equality.</summary>
    public static bool operator ==(ExcelRangeAddress left, ExcelRangeAddress right) => left.Equals(right);

    /// <summary>Compares two ExcelRangeAddress values for inequality.</summary>
    public static bool operator !=(ExcelRangeAddress left, ExcelRangeAddress right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString() => Address;
}
