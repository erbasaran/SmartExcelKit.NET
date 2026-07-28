using System.Collections;
using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Core;

/// <summary>
/// Represents a named range scoped to a workbook or a specific worksheet.
/// </summary>
public sealed class ExcelNamedRange
{
    private string _name;

    /// <summary>
    /// Gets the unique name of the range.
    /// </summary>
    public string Name
    {
        get => _name;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Named range name cannot be null or empty.", nameof(value));
            _name = value;
        }
    }

    /// <summary>
    /// Gets or sets the target range formula or address reference string (e.g., "Sheet1!$A$1:$C$10").
    /// </summary>
    public string RefersTo { get; set; }

    /// <summary>
    /// Gets or sets the optional comment attached to the named range.
    /// </summary>
    public string? Comment { get; set; }

    /// <summary>
    /// Gets the optional worksheet scope name (null if workbook-scoped).
    /// </summary>
    public string? ScopeSheetName { get; }

    /// <summary>
    /// Gets a value indicating whether this named range is scoped to a specific worksheet.
    /// </summary>
    public bool IsWorksheetScoped => ScopeSheetName != null;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelNamedRange"/> class.
    /// </summary>
    public ExcelNamedRange(string name, string refersTo, string? scopeSheetName = null)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Name cannot be null or empty.", nameof(name));
        if (string.IsNullOrWhiteSpace(refersTo)) throw new ArgumentException("RefersTo reference cannot be null or empty.", nameof(refersTo));

        _name = name;
        RefersTo = refersTo;
        ScopeSheetName = scopeSheetName;
    }
}

/// <summary>
/// Manages named ranges scoped to a workbook or worksheet.
/// </summary>
public sealed class ExcelNamedRangeCollection : IEnumerable<ExcelNamedRange>
{
    private readonly List<ExcelNamedRange> _namedRanges = [];
    private readonly string? _scopeSheetName;

    /// <summary>
    /// Gets the number of named ranges.
    /// </summary>
    public int Count => _namedRanges.Count;

    /// <summary>
    /// Gets a named range by 0-based index.
    /// </summary>
    public ExcelNamedRange this[int index] => _namedRanges[index];

    /// <summary>
    /// Gets a named range by name (case-insensitive).
    /// </summary>
    public ExcelNamedRange this[string name]
    {
        get
        {
            var nr = _namedRanges.Find(r => string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase));
            if (nr == null)
                throw new WorksheetException($"Named range '{name}' was not found.", "NAMED_RANGE_NOT_FOUND");
            return nr;
        }
    }

    internal ExcelNamedRangeCollection(string? scopeSheetName = null)
    {
        _scopeSheetName = scopeSheetName;
    }

    /// <summary>
    /// Adds a new named range.
    /// </summary>
    public ExcelNamedRange Add(string name, string refersTo, string? comment = null)
    {
        if (_namedRanges.Exists(r => string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase)))
        {
            throw new WorksheetException($"A named range with the name '{name}' already exists in this scope.", "DUPLICATE_NAMED_RANGE");
        }

        var nr = new ExcelNamedRange(name, refersTo, _scopeSheetName)
        {
            Comment = comment
        };
        _namedRanges.Add(nr);
        return nr;
    }

    /// <summary>
    /// Removes a named range by name.
    /// </summary>
    public bool Remove(string name)
    {
        int idx = _namedRanges.FindIndex(r => string.Equals(r.Name, name, StringComparison.OrdinalIgnoreCase));
        if (idx >= 0)
        {
            _namedRanges.RemoveAt(idx);
            return true;
        }
        return false;
    }

    /// <summary>
    /// Clears all named ranges in this collection.
    /// </summary>
    public void Clear() => _namedRanges.Clear();

    /// <inheritdoc />
    public IEnumerator<ExcelNamedRange> GetEnumerator() => _namedRanges.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
