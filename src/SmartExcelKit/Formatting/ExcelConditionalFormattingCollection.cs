using System.Collections;
using SmartExcelKit.Core;

namespace SmartExcelKit.Formatting;

/// <summary>
/// Manages conditional formatting rules in a worksheet.
/// </summary>
public sealed class ExcelConditionalFormattingCollection : IEnumerable<ExcelConditionalFormattingRule>
{
    private readonly ExcelWorksheet _worksheet;
    private readonly List<ExcelConditionalFormattingRule> _rules = [];

    /// <summary>
    /// Gets the count of registered rules.
    /// </summary>
    public int Count => _rules.Count;

    /// <summary>
    /// Gets the rule at the specified index.
    /// </summary>
    public ExcelConditionalFormattingRule this[int index] => _rules[index];

    internal ExcelConditionalFormattingCollection(ExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    /// <summary>
    /// Adds a cell value conditional formatting rule.
    /// </summary>
    public ExcelConditionalFormattingRule AddCellValueRule(string rangeAddress, ConditionalFormattingOperator op, string value, Styles.ExcelStyle style)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var rule = new ExcelConditionalFormattingRule(range, ConditionalFormattingType.CellValue)
        {
            Operator = op,
            Formula1 = value,
            Style = style
        };
        _rules.Add(rule);
        return rule;
    }

    /// <summary>
    /// Adds a formula conditional formatting rule.
    /// </summary>
    public ExcelConditionalFormattingRule AddFormulaRule(string rangeAddress, string formula, Styles.ExcelStyle style)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var rule = new ExcelConditionalFormattingRule(range, ConditionalFormattingType.Formula)
        {
            Formula1 = formula.TrimStart('='),
            Style = style
        };
        _rules.Add(rule);
        return rule;
    }

    /// <summary>
    /// Adds a 2-color scale conditional formatting rule.
    /// </summary>
    public ExcelConditionalFormattingRule AddColorScale2(string rangeAddress, string minColorHex, string maxColorHex)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var rule = new ExcelConditionalFormattingRule(range, ConditionalFormattingType.ColorScale)
        {
            Color1 = minColorHex,
            Color2 = maxColorHex
        };
        _rules.Add(rule);
        return rule;
    }

    /// <summary>
    /// Adds a 3-color scale conditional formatting rule.
    /// </summary>
    public ExcelConditionalFormattingRule AddThreeColorScale(string rangeAddress, string minColorHex, string midColorHex, string maxColorHex)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var rule = new ExcelConditionalFormattingRule(range, ConditionalFormattingType.ColorScale)
        {
            Color1 = minColorHex,
            Color2 = midColorHex,
            Color3 = maxColorHex
        };
        _rules.Add(rule);
        return rule;
    }

    /// <summary>
    /// Adds a data bar conditional formatting rule.
    /// </summary>
    public ExcelConditionalFormattingRule AddDataBar(string rangeAddress, string colorHex)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var rule = new ExcelConditionalFormattingRule(range, ConditionalFormattingType.DataBar)
        {
            Color1 = colorHex
        };
        _rules.Add(rule);
        return rule;
    }

    /// <summary>
    /// Clears all conditional formatting rules.
    /// </summary>
    public void Clear() => _rules.Clear();

    /// <inheritdoc />
    public IEnumerator<ExcelConditionalFormattingRule> GetEnumerator() => _rules.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
