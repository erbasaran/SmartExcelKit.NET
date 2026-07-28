using SmartExcelKit.Core;
using SmartExcelKit.Styles;

namespace SmartExcelKit.Formatting;

/// <summary>
/// Specifies the type of conditional formatting rule.
/// </summary>
public enum ConditionalFormattingType
{
    /// <summary>
    /// Evaluates cell value against a operator/value threshold.
    /// </summary>
    CellValue,

    /// <summary>
    /// Evaluates a custom Excel formula.
    /// </summary>
    Formula,

    /// <summary>
    /// Displays a 2-color or 3-color gradient scale.
    /// </summary>
    ColorScale,

    /// <summary>
    /// Displays a data bar proportional to cell value.
    /// </summary>
    DataBar,

    /// <summary>
    /// Displays an icon set based on cell value ranges.
    /// </summary>
    IconSet,

    /// <summary>
    /// Highlights top or bottom N values/percentages.
    /// </summary>
    TopBottom,

    /// <summary>
    /// Highlights duplicate values.
    /// </summary>
    DuplicateValues,

    /// <summary>
    /// Highlights unique values.
    /// </summary>
    UniqueValues
}

/// <summary>
/// Specifies conditional formatting comparison operator.
/// </summary>
public enum ConditionalFormattingOperator
{
    /// <summary>
    /// Cell value equals expected value.
    /// </summary>
    Equal,

    /// <summary>
    /// Cell value does not equal value.
    /// </summary>
    NotEqual,

    /// <summary>
    /// Cell value is greater than expected value.
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Cell value is greater than or equal to expected value.
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// Cell value is less than expected value.
    /// </summary>
    LessThan,

    /// <summary>
    /// Cell value is less than or equal to expected value.
    /// </summary>
    LessThanOrEqual,

    /// <summary>
    /// Cell value is between lower and upper bounds.
    /// </summary>
    Between,

    /// <summary>
    /// Cell value is not between bounds.
    /// </summary>
    NotBetween,

    /// <summary>
    /// Cell value contains specific text.
    /// </summary>
    ContainsText,

    /// <summary>
    /// Cell value starts with text.
    /// </summary>
    StartsWith,

    /// <summary>
    /// Cell value ends with text.
    /// </summary>
    EndsWith
}

/// <summary>
/// Represents a single conditional formatting rule applied to ranges.
/// </summary>
public sealed class ExcelConditionalFormattingRule
{
    /// <summary>
    /// Gets the target range address.
    /// </summary>
    public ExcelRangeAddress Range { get; }

    /// <summary>
    /// Gets the rule type.
    /// </summary>
    public ConditionalFormattingType RuleType { get; }

    /// <summary>
    /// Gets or sets the comparison operator for cell value rules.
    /// </summary>
    public ConditionalFormattingOperator Operator { get; set; } = ConditionalFormattingOperator.Equal;

    /// <summary>
    /// Gets or sets the primary formula or value expression.
    /// </summary>
    public string? Formula1 { get; set; }

    /// <summary>
    /// Gets or sets the secondary formula or value expression for range rules (e.g. Between).
    /// </summary>
    public string? Formula2 { get; set; }

    /// <summary>
    /// Gets or sets the custom cell style applied when condition is satisfied.
    /// </summary>
    public ExcelStyle? Style { get; set; }

    /// <summary>
    /// Gets or sets the color for DataBar or 2-color scale (HEX, e.g. "FF0000").
    /// </summary>
    public string? Color1 { get; set; }

    /// <summary>
    /// Gets or sets the secondary color for 2/3-color scales.
    /// </summary>
    public string? Color2 { get; set; }

    /// <summary>
    /// Gets or sets the tertiary color for 3-color scales.
    /// </summary>
    public string? Color3 { get; set; }

    /// <summary>
    /// Gets or sets whether further conditional formatting evaluation stops if this rule matches.
    /// </summary>
    public bool StopIfTrue { get; set; } = false;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelConditionalFormattingRule"/> class.
    /// </summary>
    public ExcelConditionalFormattingRule(ExcelRangeAddress range, ConditionalFormattingType ruleType)
    {
        Range = range;
        RuleType = ruleType;
    }
}
