using SmartExcelKit.Core;

namespace SmartExcelKit.Drawings;

/// <summary>
/// Specifies the aggregation function for pivot table data fields.
/// </summary>
public enum PivotSummaryFunction
{
    /// <summary>
    /// Sum aggregation.
    /// </summary>
    Sum,

    /// <summary>
    /// Count aggregation.
    /// </summary>
    Count,

    /// <summary>
    /// Average aggregation.
    /// </summary>
    Average,

    /// <summary>
    /// Minimum value.
    /// </summary>
    Min,

    /// <summary>
    /// Maximum value.
    /// </summary>
    Max,

    /// <summary>
    /// Standard deviation.
    /// </summary>
    StdDev,

    /// <summary>
    /// Variance.
    /// </summary>
    Var
}

/// <summary>
/// Represents a field in a pivot table layout.
/// </summary>
public sealed class PivotField
{
    /// <summary>
    /// Gets the field column name.
    /// </summary>
    public string Name { get; }

    /// <summary>
    /// Gets or sets the custom display header for data fields.
    /// </summary>
    public string? CustomName { get; set; }

    /// <summary>
    /// Gets or sets the aggregation function for data fields.
    /// </summary>
    public PivotSummaryFunction Function { get; set; } = PivotSummaryFunction.Sum;

    /// <summary>
    /// Initializes a new instance of the <see cref="PivotField"/> class.
    /// </summary>
    public PivotField(string name, PivotSummaryFunction function = PivotSummaryFunction.Sum)
    {
        Name = name;
        Function = function;
    }
}

/// <summary>
/// Represents a Pivot Table in a worksheet.
/// </summary>
public sealed class ExcelPivotTable
{
    private readonly List<PivotField> _rowFields = [];
    private readonly List<PivotField> _columnFields = [];
    private readonly List<PivotField> _dataFields = [];
    private readonly List<PivotField> _filterFields = [];

    /// <summary>
    /// Gets the unique name of the pivot table.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets or sets the source data range address (e.g. "Sheet1!A1:D100").
    /// </summary>
    public string SourceRange { get; set; }

    /// <summary>
    /// Gets or sets the target top-left cell location for the pivot table.
    /// </summary>
    public CellAddress TargetCell { get; set; }

    /// <summary>
    /// Gets the list of row fields.
    /// </summary>
    public IReadOnlyList<PivotField> RowFields => _rowFields;

    /// <summary>
    /// Gets the list of column fields.
    /// </summary>
    public IReadOnlyList<PivotField> ColumnFields => _columnFields;

    /// <summary>
    /// Gets the list of data (value) fields.
    /// </summary>
    public IReadOnlyList<PivotField> DataFields => _dataFields;

    /// <summary>
    /// Gets the list of page filter fields.
    /// </summary>
    public IReadOnlyList<PivotField> FilterFields => _filterFields;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelPivotTable"/> class.
    /// </summary>
    public ExcelPivotTable(string name, string sourceRange, CellAddress targetCell)
    {
        Name = name;
        SourceRange = sourceRange;
        TargetCell = targetCell;
    }

    /// <summary>
    /// Adds a row field to the pivot table.
    /// </summary>
    public PivotField AddRowField(string fieldName)
    {
        var field = new PivotField(fieldName);
        _rowFields.Add(field);
        return field;
    }

    /// <summary>
    /// Adds a column field to the pivot table.
    /// </summary>
    public PivotField AddColumnField(string fieldName)
    {
        var field = new PivotField(fieldName);
        _columnFields.Add(field);
        return field;
    }

    /// <summary>
    /// Adds a data (value) field to the pivot table.
    /// </summary>
    public PivotField AddDataField(string fieldName, PivotSummaryFunction function = PivotSummaryFunction.Sum, string? customName = null)
    {
        var field = new PivotField(fieldName, function)
        {
            CustomName = customName ?? $"{function} of {fieldName}"
        };
        _dataFields.Add(field);
        return field;
    }

    /// <summary>
    /// Adds a filter (page) field to the pivot table.
    /// </summary>
    public PivotField AddFilterField(string fieldName)
    {
        var field = new PivotField(fieldName);
        _filterFields.Add(field);
        return field;
    }
}
