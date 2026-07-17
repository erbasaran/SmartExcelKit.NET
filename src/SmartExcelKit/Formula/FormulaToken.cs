namespace SmartExcelKit.Formula;

/// <summary>
/// Specifies the type of a formula token.
/// </summary>
internal enum FormulaTokenType
{
    Number,
    String,
    Boolean,
    Reference,
    Identifier,
    Operator,
    LeftParenthesis,
    RightParenthesis,
    Comma,
    Error
}

/// <summary>
/// Represents a token in an Excel formula.
/// </summary>
internal sealed class FormulaToken
{
    public FormulaTokenType Type { get; }
    public string Value { get; }

    public FormulaToken(FormulaTokenType type, string value)
    {
        Type = type;
        Value = value;
    }

    public override string ToString() => $"{Type}: {Value}";
}
