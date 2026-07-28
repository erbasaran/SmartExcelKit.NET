namespace SmartExcelKit.Drawings;

/// <summary>
/// Specifies the type of Excel chart.
/// </summary>
public enum ChartType
{
    /// <summary>
    /// Vertical Column chart.
    /// </summary>
    Column,

    /// <summary>
    /// Horizontal Bar chart.
    /// </summary>
    Bar,

    /// <summary>
    /// Line chart.
    /// </summary>
    Line,

    /// <summary>
    /// Pie chart.
    /// </summary>
    Pie,

    /// <summary>
    /// Doughnut chart.
    /// </summary>
    Doughnut,

    /// <summary>
    /// Area chart.
    /// </summary>
    Area,

    /// <summary>
    /// Scatter plot chart.
    /// </summary>
    Scatter
}

/// <summary>
/// Specifies chart legend positions.
/// </summary>
public enum LegendPosition
{
    /// <summary>
    /// Legend at top.
    /// </summary>
    Top,

    /// <summary>
    /// Legend at bottom.
    /// </summary>
    Bottom,

    /// <summary>
    /// Legend at left.
    /// </summary>
    Left,

    /// <summary>
    /// Legend at right.
    /// </summary>
    Right,

    /// <summary>
    /// Legend in top right corner.
    /// </summary>
    TopRight,

    /// <summary>
    /// Legend hidden.
    /// </summary>
    None
}

/// <summary>
/// Represents a chart series data binding.
/// </summary>
public sealed class ChartSeries
{
    /// <summary>
    /// Gets or sets the series title or header cell reference.
    /// </summary>
    public string Name { get; set; }

    /// <summary>
    /// Gets or sets the range reference containing Y values (e.g. "Sheet1!B2:B10").
    /// </summary>
    public string ValuesRange { get; set; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ChartSeries"/> class.
    /// </summary>
    public ChartSeries(string name, string valuesRange)
    {
        Name = name;
        ValuesRange = valuesRange;
    }
}

/// <summary>
/// Represents a chart embedded in a worksheet.
/// </summary>
public sealed class ExcelChart
{
    private readonly List<ChartSeries> _series = [];

    /// <summary>
    /// Gets the unique identifier of the chart.
    /// </summary>
    public string Id { get; } = Guid.NewGuid().ToString("N");

    /// <summary>
    /// Gets or sets the chart title.
    /// </summary>
    public string Title { get; set; } = "Chart Title";

    /// <summary>
    /// Gets or sets the chart type.
    /// </summary>
    public ChartType ChartType { get; set; } = ChartType.Column;

    /// <summary>
    /// Gets or sets the legend position.
    /// </summary>
    public LegendPosition LegendPosition { get; set; } = LegendPosition.Right;

    /// <summary>
    /// Gets or sets the range reference for category axis labels (X values, e.g. "Sheet1!A2:A10").
    /// </summary>
    public string? CategoryRange { get; set; }

    /// <summary>
    /// Gets the list of chart series.
    /// </summary>
    public IReadOnlyList<ChartSeries> Series => _series;

    /// <summary>
    /// Gets or sets the top 1-based row anchor.
    /// </summary>
    public int TopRow { get; set; }

    /// <summary>
    /// Gets or sets the left 1-based column anchor.
    /// </summary>
    public int LeftColumn { get; set; }

    /// <summary>
    /// Gets or sets chart width in pixels.
    /// </summary>
    public int Width { get; set; } = 480;

    /// <summary>
    /// Gets or sets chart height in pixels.
    /// </summary>
    public int Height { get; set; } = 320;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelChart"/> class.
    /// </summary>
    public ExcelChart(ChartType chartType, int topRow, int leftColumn, int width = 480, int height = 320)
    {
        ChartType = chartType;
        TopRow = topRow;
        LeftColumn = leftColumn;
        Width = width;
        Height = height;
    }

    /// <summary>
    /// Adds a data series to the chart.
    /// </summary>
    public ChartSeries AddSeries(string seriesName, string valuesRange)
    {
        var s = new ChartSeries(seriesName, valuesRange);
        _series.Add(s);
        return s;
    }
}
