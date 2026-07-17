using System.Text;

namespace SmartExcelKit.Core;

/// <summary>
/// Represents a 1-based row and column address in a worksheet.
/// </summary>
public readonly struct CellAddress : IEquatable<CellAddress>
{
    /// <summary>
    /// Gets the 1-based row number.
    /// </summary>
    public int Row { get; }

    /// <summary>
    /// Gets the 1-based column number.
    /// </summary>
    public int Column { get; }

    /// <summary>
    /// Gets the address in A1 notation (e.g., "A1", "C10").
    /// </summary>
    public string Address => $"{GetColumnName(Column)}{Row}";

    /// <summary>
    /// Initializes a new instance of the <see cref="CellAddress"/> struct.
    /// </summary>
    /// <param name="row">The 1-based row number.</param>
    /// <param name="column">The 1-based column number.</param>
    /// <exception cref="ArgumentOutOfRangeException">Thrown if row or column is less than 1.</exception>
    public CellAddress(int row, int column)
    {
        if (row < 1) throw new ArgumentOutOfRangeException(nameof(row), "Row must be greater than or equal to 1.");
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column), "Column must be greater than or equal to 1.");
        Row = row;
        Column = column;
    }

    /// <summary>
    /// Parses an A1-style cell reference (e.g. "B12", "$A$1") into a CellAddress.
    /// </summary>
    /// <param name="address">The A1 cell reference.</param>
    /// <returns>A new <see cref="CellAddress"/> instance.</returns>
    public static CellAddress Parse(string address)
    {
        if (string.IsNullOrWhiteSpace(address))
            throw new ArgumentException("Address cannot be null or empty.", nameof(address));

        int colIndex = 0;
        int rowIndex = 0;

        int i = 0;
        // Skip absolute identifiers
        if (address[i] == '$') i++;

        while (i < address.Length && char.IsLetter(address[i]))
        {
            colIndex = colIndex * 26 + (char.ToUpperInvariant(address[i]) - 'A' + 1);
            i++;
        }

        if (i < address.Length && address[i] == '$') i++;

        while (i < address.Length && char.IsDigit(address[i]))
        {
            rowIndex = rowIndex * 10 + (address[i] - '0');
            i++;
        }

        if (colIndex == 0 || rowIndex == 0)
            throw new FormatException($"Invalid cell address format: '{address}'.");

        return new CellAddress(rowIndex, colIndex);
    }

    /// <summary>
    /// Converts a 1-based column number into its alphabetical name (e.g., 1 -> "A", 28 -> "AB").
    /// </summary>
    /// <param name="column">The 1-based column number.</param>
    /// <returns>The alphabetical column name.</returns>
    public static string GetColumnName(int column)
    {
        if (column < 1) throw new ArgumentOutOfRangeException(nameof(column), "Column must be >= 1.");

        var sb = new StringBuilder();
        int temp = column;
        while (temp > 0)
        {
            int modulo = (temp - 1) % 26;
            sb.Insert(0, (char)('A' + modulo));
            temp = (temp - modulo) / 26;
        }
        return sb.ToString();
    }

    /// <summary>
    /// Converts an alphabetical column name (e.g., "A", "AB") into its 1-based number.
    /// </summary>
    /// <param name="columnName">The column name.</param>
    /// <returns>The 1-based column number.</returns>
    public static int GetColumnNumber(string columnName)
    {
        if (string.IsNullOrWhiteSpace(columnName))
            throw new ArgumentException("Column name cannot be null or empty.", nameof(columnName));

        int col = 0;
        foreach (char c in columnName)
        {
            if (!char.IsLetter(c))
                throw new ArgumentException($"Invalid character '{c}' in column name.", nameof(columnName));
            col = col * 26 + (char.ToUpperInvariant(c) - 'A' + 1);
        }
        return col;
    }

    /// <inheritdoc />
    public bool Equals(CellAddress other) => Row == other.Row && Column == other.Column;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is CellAddress other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode()
    {
        // Simple hash combining row and column
        return (Row * 397) ^ Column;
    }

    /// <summary>Compares two CellAddress values for equality.</summary>
    public static bool operator ==(CellAddress left, CellAddress right) => left.Equals(right);

    /// <summary>Compares two CellAddress values for inequality.</summary>
    public static bool operator !=(CellAddress left, CellAddress right) => !left.Equals(right);

    /// <inheritdoc />
    public override string ToString() => Address;
}
