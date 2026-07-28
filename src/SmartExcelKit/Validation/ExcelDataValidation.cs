using SmartExcelKit.Core;

namespace SmartExcelKit.Validation;

/// <summary>
/// Specifies the type of data validation.
/// </summary>
public enum ValidationType
{
    /// <summary>
    /// Any value is allowed (validation disabled).
    /// </summary>
    Any,

    /// <summary>
    /// Restricts input to whole numbers.
    /// </summary>
    WholeNumber,

    /// <summary>
    /// Restricts input to decimal values.
    /// </summary>
    Decimal,

    /// <summary>
    /// Restricts input to a defined explicit list or formula reference.
    /// </summary>
    List,

    /// <summary>
    /// Restricts input to date values.
    /// </summary>
    Date,

    /// <summary>
    /// Restricts input to time values.
    /// </summary>
    Time,

    /// <summary>
    /// Restricts input string length.
    /// </summary>
    TextLength,

    /// <summary>
    /// Validates using a custom boolean formula.
    /// </summary>
    Custom
}

/// <summary>
/// Specifies the validation operator.
/// </summary>
public enum ValidationOperator
{
    /// <summary>
    /// Value is between formula1 and formula2.
    /// </summary>
    Between,

    /// <summary>
    /// Value is not between formula1 and formula2.
    /// </summary>
    NotBetween,

    /// <summary>
    /// Value equals formula1.
    /// </summary>
    Equal,

    /// <summary>
    /// Value does not equal formula1.
    /// </summary>
    NotEqual,

    /// <summary>
    /// Value is greater than formula1.
    /// </summary>
    GreaterThan,

    /// <summary>
    /// Value is less than formula1.
    /// </summary>
    LessThan,

    /// <summary>
    /// Value is greater than or equal to formula1.
    /// </summary>
    GreaterThanOrEqual,

    /// <summary>
    /// Value is less than or equal to formula1.
    /// </summary>
    LessThanOrEqual
}

/// <summary>
/// Specifies the error alert style when invalid input is entered.
/// </summary>
public enum ValidationErrorStyle
{
    /// <summary>
    /// Stops user from entering invalid data (Stop modal).
    /// </summary>
    Stop,

    /// <summary>
    /// Displays a warning dialog but allows user to proceed.
    /// </summary>
    Warning,

    /// <summary>
    /// Displays an informational message and allows user to proceed.
    /// </summary>
    Information
}

/// <summary>
/// Represents a Data Validation rule applied to a range in a worksheet.
/// </summary>
public sealed class ExcelDataValidation
{
    /// <summary>
    /// Gets the target cell range address.
    /// </summary>
    public ExcelRangeAddress Range { get; }

    /// <summary>
    /// Gets the validation type.
    /// </summary>
    public ValidationType ValidationType { get; }

    /// <summary>
    /// Gets or sets the validation operator.
    /// </summary>
    public ValidationOperator Operator { get; set; } = ValidationOperator.Between;

    /// <summary>
    /// Gets or sets the primary constraint, list string, or formula.
    /// </summary>
    public string? Formula1 { get; set; }

    /// <summary>
    /// Gets or sets the secondary constraint formula (for Between/NotBetween).
    /// </summary>
    public string? Formula2 { get; set; }

    /// <summary>
    /// Gets or sets whether empty cells pass validation.
    /// </summary>
    public bool AllowBlank { get; set; } = true;

    /// <summary>
    /// Gets or sets whether an in-cell drop-down list is displayed for List validations.
    /// </summary>
    public bool ShowDropDown { get; set; } = true;

    /// <summary>
    /// Gets or sets whether an input message prompt is shown when cell is selected.
    /// </summary>
    public bool ShowInputMessage { get; set; } = true;

    /// <summary>
    /// Gets or sets the input prompt title.
    /// </summary>
    public string? PromptTitle { get; set; }

    /// <summary>
    /// Gets or sets the input prompt message body.
    /// </summary>
    public string? Prompt { get; set; }

    /// <summary>
    /// Gets or sets whether an error alert dialog is shown on invalid input.
    /// </summary>
    public bool ShowErrorMessage { get; set; } = true;

    /// <summary>
    /// Gets or sets the error alert style.
    /// </summary>
    public ValidationErrorStyle ErrorStyle { get; set; } = ValidationErrorStyle.Stop;

    /// <summary>
    /// Gets or sets the error dialog title.
    /// </summary>
    public string? ErrorTitle { get; set; }

    /// <summary>
    /// Gets or sets the error dialog message text (alias for <see cref="Error"/>).
    /// </summary>
    public string? ErrorMessage { get => Error; set => Error = value; }

    /// <summary>
    /// Gets or sets the error dialog message text.
    /// </summary>
    public string? Error { get; set; }

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelDataValidation"/> class.
    /// </summary>
    public ExcelDataValidation(ExcelRangeAddress range, ValidationType validationType)
    {
        Range = range;
        ValidationType = validationType;
    }
}
