using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Formula;

/// <summary>
/// Evaluates parsed AST formula nodes against a worksheet context.
/// </summary>
public static class FormulaEvaluator
{
    /// <summary>
    /// Evaluates a formula string in the context of an active worksheet and a target cell.
    /// </summary>
    public static object? Evaluate(string formula, ExcelWorksheet activeWorksheet, CellAddress currentCell)
    {
        if (string.IsNullOrWhiteSpace(formula)) return null;

        var ast = FormulaParser.Parse(formula);
        return EvaluateNode(ast, activeWorksheet);
    }

    /// <summary>
    /// Evaluates an AST Node internally.
    /// </summary>
    internal static object? EvaluateNode(FormulaNode node, ExcelWorksheet sheet)
    {
        if (node == null) throw new ArgumentNullException(nameof(node));

        switch (node)
        {
            case LiteralNode literal:
                return literal.Value;

            case ReferenceNode reference:
                var targetSheet = string.IsNullOrEmpty(reference.SheetName)
                    ? sheet
                    : (sheet.Workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, reference.SheetName, StringComparison.OrdinalIgnoreCase))
                       ?? throw new FormulaException($"Worksheet '{reference.SheetName}' was not found.", "WORKSHEET_NOT_FOUND"));

                var addr = reference.RangeAddress;

                if (addr.StartRow == addr.EndRow && addr.StartColumn == addr.EndColumn)
                {
                    // Single Cell
                    var cellAddr = new CellAddress(addr.StartRow, addr.StartColumn);
                    return GetCellValueEvaluated(targetSheet, cellAddr);
                }
                else
                {
                    // Range of cells: return a flat list of evaluated cell values
                    var list = new List<object?>();
                    for (int r = addr.StartRow; r <= addr.EndRow; r++)
                    {
                        for (int c = addr.StartColumn; c <= addr.EndColumn; c++)
                        {
                            list.Add(GetCellValueEvaluated(targetSheet, new CellAddress(r, c)));
                        }
                    }
                    return list;
                }

            case OperatorNode opNode:
                var leftVal = EvaluateNode(opNode.Left, sheet);
                var rightVal = opNode.Right != null ? EvaluateNode(opNode.Right, sheet) : null;

                return EvaluateOperator(opNode.Operator, leftVal, rightVal);

            case FunctionNode funcNode:
                var evaluatedArgs = funcNode.Arguments.Select(arg => EvaluateNode(arg, sheet)).ToList();
                return EvaluateFunction(funcNode.FunctionName, evaluatedArgs);

            default:
                throw new FormulaException($"Unsupported AST node type: '{node.GetType().Name}'", "UNSUPPORTED_NODE");
        }
    }

    private static object? GetCellValueEvaluated(ExcelWorksheet sheet, CellAddress address)
    {
        var formula = sheet.GetCellFormula(address);
        if (!string.IsNullOrEmpty(formula))
        {
            // Evaluate the formula recursively
            return Evaluate(formula!, sheet, address);
        }
        return sheet.GetCellValue(address);
    }

    private static object? EvaluateOperator(string op, object? left, object? right)
    {
        if (right == null)
        {
            // Unary operators
            if (op == "-") return -ConvertToDouble(left);
            if (op == "+") return ConvertToDouble(left);
            throw new FormulaException($"Unsupported unary operator: '{op}'", "INVALID_OPERATOR");
        }

        // Binary operators
        if (op == "&")
        {
            return (left?.ToString() ?? string.Empty) + (right?.ToString() ?? string.Empty);
        }

        if (op == "+" || op == "-" || op == "*" || op == "/" || op == "^")
        {
            double l = ConvertToDouble(left);
            double r = ConvertToDouble(right);
            return op switch
            {
                "+" => l + r,
                "-" => l - r,
                "*" => l * r,
                "/" => r == 0.0 ? throw new FormulaException("Division by zero.", "DIVISION_BY_ZERO") : l / r,
                "^" => Math.Pow(l, r),
                _ => throw new FormulaException($"Invalid arithmetic operator '{op}'", "INVALID_OPERATOR")
            };
        }

        // Comparisons
        if (op == "=" || op == "<>" || op == "<" || op == ">" || op == "<=" || op == ">=")
        {
            int cmp = CompareValues(left, right);
            return op switch
            {
                "=" => cmp == 0,
                "<>" => cmp != 0,
                "<" => cmp < 0,
                ">" => cmp > 0,
                "<=" => cmp <= 0,
                ">=" => cmp >= 0,
                _ => throw new FormulaException($"Invalid comparison operator '{op}'", "INVALID_OPERATOR")
            };
        }

        throw new FormulaException($"Unsupported operator: '{op}'", "INVALID_OPERATOR");
    }

    private static object? EvaluateFunction(string name, List<object?> args)
    {
        switch (name.ToUpperInvariant())
        {
            case "SUM":
                return Flatten(args).Select(ConvertToDouble).Sum();

            case "AVERAGE":
                var sumList = Flatten(args).Select(ConvertToDouble).ToList();
                return sumList.Count == 0 ? 0.0 : sumList.Average();

            case "COUNT":
                return Flatten(args).Count(val => val is double || val is float || val is int || val is decimal || (val is string s && double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out _)));

            case "COUNTA":
                return Flatten(args).Count(val => val != null && !string.IsNullOrEmpty(val.ToString()));

            case "MIN":
                var minList = Flatten(args).Select(ConvertToDouble).ToList();
                return minList.Count == 0 ? 0.0 : minList.Min();

            case "MAX":
                var maxList = Flatten(args).Select(ConvertToDouble).ToList();
                return maxList.Count == 0 ? 0.0 : maxList.Max();

            case "IF":
                if (args.Count < 2) throw new FormulaException("IF function requires at least 2 arguments.", "ARGUMENT_COUNT_ERROR");
                bool condition = ConvertToBoolean(args[0]);
                if (condition) return args[1];
                return args.Count > 2 ? args[2] : null;

            case "AND":
                if (args.Count == 0) return false;
                return args.All(ConvertToBoolean);

            case "OR":
                if (args.Count == 0) return false;
                return args.Any(ConvertToBoolean);

            case "NOT":
                if (args.Count != 1) throw new FormulaException("NOT function requires exactly 1 argument.", "ARGUMENT_COUNT_ERROR");
                return !ConvertToBoolean(args[0]);

            case "LEFT":
                if (args.Count == 0) return string.Empty;
                string leftStr = args[0]?.ToString() ?? string.Empty;
                int leftLen = args.Count > 1 ? (int)ConvertToDouble(args[1]) : 1;
                if (leftLen < 0) leftLen = 0;
                return leftLen >= leftStr.Length ? leftStr : leftStr.Substring(0, leftLen);

            case "RIGHT":
                if (args.Count == 0) return string.Empty;
                string rightStr = args[0]?.ToString() ?? string.Empty;
                int rightLen = args.Count > 1 ? (int)ConvertToDouble(args[1]) : 1;
                if (rightLen < 0) rightLen = 0;
                return rightLen >= rightStr.Length ? rightStr : rightStr.Substring(rightStr.Length - rightLen, rightLen);

            case "LEN":
                if (args.Count == 0) return 0;
                return args[0]?.ToString()?.Length ?? 0;

            case "CONCAT":
                return string.Concat(Flatten(args).Select(a => a?.ToString() ?? string.Empty));

            case "TODAY":
                return DateTime.Today;

            case "NOW":
                return DateTime.Now;

            default:
                throw new FormulaException($"Function '{name}' is not supported yet by the evaluator engine.", "UNSUPPORTED_FUNCTION");
        }
    }

    private static double ConvertToDouble(object? val)
    {
        if (val == null) return 0.0;
        if (val is double d) return d;
        if (val is float f) return f;
        if (val is int i) return i;
        if (val is decimal dec) return (double)dec;
        if (val is string s && double.TryParse(s, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double parsed)) return parsed;
        if (val is bool b) return b ? 1.0 : 0.0;
        return 0.0;
    }

    private static bool ConvertToBoolean(object? val)
    {
        if (val == null) return false;
        if (val is bool b) return b;
        if (val is double d) return d != 0.0;
        if (val is int i) return i != 0;
        if (val is string s && bool.TryParse(s, out bool parsed)) return parsed;
        return false;
    }

    private static int CompareValues(object? left, object? right)
    {
        if (left == null && right == null) return 0;
        if (left == null) return -1;
        if (right == null) return 1;

        if (left is double || left is float || left is int || right is double || right is float || right is int)
        {
            double l = ConvertToDouble(left);
            double r = ConvertToDouble(right);
            return l.CompareTo(r);
        }

        if (left is bool lBool && right is bool rBool)
        {
            return lBool.CompareTo(rBool);
        }

        return string.Compare(left.ToString(), right.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    private static IEnumerable<object?> Flatten(List<object?> args)
    {
        foreach (var arg in args)
        {
            if (arg is IEnumerable<object?> list && !(arg is string))
            {
                foreach (var inner in list)
                {
                    yield return inner;
                }
            }
            else
            {
                yield return arg;
            }
        }
    }
}
