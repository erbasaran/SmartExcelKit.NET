using SmartExcelKit.Core;

namespace SmartExcelKit.Formula;

/// <summary>
/// Base class for all nodes in the formula AST.
/// </summary>
internal abstract class FormulaNode
{
}

/// <summary>
/// Represents a constant value node (number, string, boolean).
/// </summary>
internal sealed class LiteralNode : FormulaNode
{
    public object? Value { get; }

    public LiteralNode(object? value)
    {
        Value = value;
    }

    public override string ToString() => Value?.ToString() ?? "null";
}

/// <summary>
/// Represents a single cell or range reference, possibly containing a worksheet qualifier (e.g. Sheet1!A1:C3).
/// </summary>
internal sealed class ReferenceNode : FormulaNode
{
    public string? SheetName { get; }
    public ExcelRangeAddress RangeAddress { get; }

    public ReferenceNode(ExcelRangeAddress rangeAddress, string? sheetName = null)
    {
        RangeAddress = rangeAddress;
        SheetName = sheetName;
    }

    public override string ToString() => string.IsNullOrEmpty(SheetName) ? RangeAddress.Address : $"{SheetName}!{RangeAddress.Address}";
}

/// <summary>
/// Represents unary or binary operations.
/// </summary>
internal sealed class OperatorNode : FormulaNode
{
    public string Operator { get; }
    public FormulaNode Left { get; }
    public FormulaNode? Right { get; } // Null for unary operator (e.g., unary minus)

    public OperatorNode(string op, FormulaNode left, FormulaNode? right = null)
    {
        Operator = op;
        Left = left;
        Right = right;
    }

    public override string ToString() => Right == null ? $"({Operator}{Left})" : $"({Left} {Operator} {Right})";
}

/// <summary>
/// Represents an Excel function call with zero or more parameter expression arguments.
/// </summary>
internal sealed class FunctionNode : FormulaNode
{
    public string FunctionName { get; }
    public IReadOnlyList<FormulaNode> Arguments { get; }

    public FunctionNode(string functionName, IReadOnlyList<FormulaNode> arguments)
    {
        FunctionName = functionName;
        Arguments = arguments;
    }

    public override string ToString() => $"{FunctionName}({string.Join(", ", Arguments)})";
}
