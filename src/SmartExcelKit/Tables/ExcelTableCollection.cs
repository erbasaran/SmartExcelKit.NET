using System.Collections;
using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Tables;

/// <summary>
/// Represents a collection of <see cref="ExcelTable"/> instances in a worksheet.
/// </summary>
public sealed class ExcelTableCollection : IEnumerable<ExcelTable>
{
    private readonly ExcelWorksheet _worksheet;
    private readonly List<ExcelTable> _tables = [];

    /// <summary>
    /// Gets the number of tables in the worksheet.
    /// </summary>
    public int Count => _tables.Count;

    /// <summary>
    /// Gets a table by 0-based index.
    /// </summary>
    public ExcelTable this[int index] => _tables[index];

    /// <summary>
    /// Gets a table by name (case-insensitive).
    /// </summary>
    public ExcelTable this[string name]
    {
        get
        {
            var table = _tables.Find(t => string.Equals(t.Name, name, StringComparison.OrdinalIgnoreCase));
            if (table == null)
                throw new WorksheetException($"Excel table '{name}' was not found.", "TABLE_NOT_FOUND");
            return table;
        }
    }

    internal ExcelTableCollection(ExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    /// <summary>
    /// Adds a new table over the specified cell range.
    /// </summary>
    /// <param name="rangeAddress">The range address (e.g. "A1:D10").</param>
    /// <param name="tableName">Optional name of the table. Defaults to "Table1", "Table2", etc.</param>
    /// <returns>The created <see cref="ExcelTable"/>.</returns>
    public ExcelTable Add(string rangeAddress, string? tableName = null)
    {
        return Add(ExcelRangeAddress.Parse(rangeAddress), tableName);
    }

    /// <summary>
    /// Adds a new table over the specified cell range.
    /// </summary>
    /// <param name="range">The range address struct.</param>
    /// <param name="tableName">Optional name of the table.</param>
    /// <returns>The created <see cref="ExcelTable"/>.</returns>
    public ExcelTable Add(ExcelRangeAddress range, string? tableName = null)
    {
        if (string.IsNullOrWhiteSpace(tableName))
        {
            tableName = $"Table{_tables.Count + 1}";
        }

        if (_tables.Exists(t => string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase)))
        {
            throw new WorksheetException($"A table named '{tableName}' already exists in the worksheet.", "DUPLICATE_TABLE_NAME");
        }

        var table = new ExcelTable(_worksheet, tableName!, range);
        _tables.Add(table);
        return table;
    }

    /// <summary>
    /// Removes a table by name.
    /// </summary>
    public bool Remove(string tableName)
    {
        int idx = _tables.FindIndex(t => string.Equals(t.Name, tableName, StringComparison.OrdinalIgnoreCase));
        if (idx >= 0)
        {
            _tables.RemoveAt(idx);
            return true;
        }
        return false;
    }

    /// <summary>
    /// Clears all tables from the worksheet.
    /// </summary>
    public void Clear() => _tables.Clear();

    /// <inheritdoc />
    public IEnumerator<ExcelTable> GetEnumerator() => _tables.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
