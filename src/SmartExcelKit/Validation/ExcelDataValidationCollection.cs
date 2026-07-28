using System.Collections;
using SmartExcelKit.Core;

namespace SmartExcelKit.Validation;

/// <summary>
/// Manages data validation rules in a worksheet.
/// </summary>
public sealed class ExcelDataValidationCollection : IEnumerable<ExcelDataValidation>
{
    private readonly ExcelWorksheet _worksheet;
    private readonly List<ExcelDataValidation> _validations = [];

    /// <summary>
    /// Gets the count of registered validation rules.
    /// </summary>
    public int Count => _validations.Count;

    /// <summary>
    /// Gets the validation rule at the specified index.
    /// </summary>
    public ExcelDataValidation this[int index] => _validations[index];

    internal ExcelDataValidationCollection(ExcelWorksheet worksheet)
    {
        _worksheet = worksheet ?? throw new ArgumentNullException(nameof(worksheet));
    }

    /// <summary>
    /// Adds a list (drop-down) validation rule using a comma-separated values string or range formula.
    /// </summary>
    public ExcelDataValidation AddListValidation(string rangeAddress, string listValuesOrFormula)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var validation = new ExcelDataValidation(range, ValidationType.List)
        {
            Formula1 = listValuesOrFormula
        };
        _validations.Add(validation);
        return validation;
    }

    /// <summary>
    /// Adds an explicit string list (drop-down) validation rule.
    /// </summary>
    public ExcelDataValidation AddListValidation(string rangeAddress, IEnumerable<string> items)
    {
        string listStr = $"\"{string.Join(",", items)}\"";
        return AddListValidation(rangeAddress, listStr);
    }

    /// <summary>
    /// Adds a whole number range validation rule.
    /// </summary>
    public ExcelDataValidation AddWholeNumberValidation(string rangeAddress, ValidationOperator op, int min, int max = 0)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var validation = new ExcelDataValidation(range, ValidationType.WholeNumber)
        {
            Operator = op,
            Formula1 = min.ToString(),
            Formula2 = op == ValidationOperator.Between || op == ValidationOperator.NotBetween ? max.ToString() : null
        };
        _validations.Add(validation);
        return validation;
    }

    /// <summary>
    /// Adds a decimal range validation rule.
    /// </summary>
    public ExcelDataValidation AddDecimalValidation(string rangeAddress, ValidationOperator op, double min, double max = 0.0)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var validation = new ExcelDataValidation(range, ValidationType.Decimal)
        {
            Operator = op,
            Formula1 = min.ToString(System.Globalization.CultureInfo.InvariantCulture),
            Formula2 = op == ValidationOperator.Between || op == ValidationOperator.NotBetween ? max.ToString(System.Globalization.CultureInfo.InvariantCulture) : null
        };
        _validations.Add(validation);
        return validation;
    }

    /// <summary>
    /// Adds a date validation rule.
    /// </summary>
    public ExcelDataValidation AddDateValidation(string rangeAddress, ValidationOperator op, DateTime minDate, DateTime? maxDate = null)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var validation = new ExcelDataValidation(range, ValidationType.Date)
        {
            Operator = op,
            Formula1 = minDate.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture),
            Formula2 = maxDate.HasValue ? maxDate.Value.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture) : null
        };
        _validations.Add(validation);
        return validation;
    }

    /// <summary>
    /// Adds a text length validation rule.
    /// </summary>
    public ExcelDataValidation AddTextLengthValidation(string rangeAddress, ValidationOperator op, int minLength, int maxLength = 0)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var validation = new ExcelDataValidation(range, ValidationType.TextLength)
        {
            Operator = op,
            Formula1 = minLength.ToString(),
            Formula2 = op == ValidationOperator.Between || op == ValidationOperator.NotBetween ? maxLength.ToString() : null
        };
        _validations.Add(validation);
        return validation;
    }

    /// <summary>
    /// Adds a custom formula validation rule.
    /// </summary>
    public ExcelDataValidation AddCustomValidation(string rangeAddress, string formula)
    {
        var range = ExcelRangeAddress.Parse(rangeAddress);
        var validation = new ExcelDataValidation(range, ValidationType.Custom)
        {
            Formula1 = formula.TrimStart('=')
        };
        _validations.Add(validation);
        return validation;
    }

    /// <summary>
    /// Clears all data validation rules.
    /// </summary>
    public void Clear() => _validations.Clear();

    /// <inheritdoc />
    public IEnumerator<ExcelDataValidation> GetEnumerator() => _validations.GetEnumerator();

    /// <inheritdoc />
    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
}
