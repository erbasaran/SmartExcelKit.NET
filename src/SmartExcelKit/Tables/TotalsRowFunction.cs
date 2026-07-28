namespace SmartExcelKit.Tables;

/// <summary>
/// Specifies the aggregation function for an Excel Table totals row cell.
/// </summary>
public enum TotalsRowFunction
{
    /// <summary>
    /// No aggregation function.
    /// </summary>
    None,

    /// <summary>
    /// Calculates the sum of values.
    /// </summary>
    Sum,

    /// <summary>
    /// Calculates the count of non-empty cells.
    /// </summary>
    Count,

    /// <summary>
    /// Calculates the count of numeric cells.
    /// </summary>
    CountNumbers,

    /// <summary>
    /// Calculates the average of numeric values.
    /// </summary>
    Average,

    /// <summary>
    /// Finds the minimum value.
    /// </summary>
    Min,

    /// <summary>
    /// Finds the maximum value.
    /// </summary>
    Max,

    /// <summary>
    /// Calculates the sample standard deviation.
    /// </summary>
    StdDev,

    /// <summary>
    /// Calculates the sample variance.
    /// </summary>
    Var,

    /// <summary>
    /// Uses a custom formula for the totals row.
    /// </summary>
    Custom
}
