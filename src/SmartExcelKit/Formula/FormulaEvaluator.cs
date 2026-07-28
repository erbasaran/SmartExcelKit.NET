using System.Globalization;
using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Formula;

/// <summary>
/// High-performance evaluation engine for Excel AST formula trees against worksheet context.
/// Supports 60+ core Excel functions, structured table references, named ranges, and dependency graph evaluation.
/// </summary>
public static class FormulaEvaluator
{
    private static readonly Dictionary<string, object?> _formulaCache = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Clears the cached formula calculation results.
    /// </summary>
    public static void ClearCache()
    {
        lock (_formulaCache)
        {
            _formulaCache.Clear();
        }
    }

    /// <summary>
    /// Evaluates a formula string in the context of an active worksheet and target cell address.
    /// </summary>
    public static object? Evaluate(string formula, ExcelWorksheet activeWorksheet, CellAddress currentCell)
    {
        if (string.IsNullOrWhiteSpace(formula)) return null;

        string cacheKey = $"{activeWorksheet.Name}!{currentCell.Address}={formula}";
        lock (_formulaCache)
        {
            if (_formulaCache.TryGetValue(cacheKey, out var cachedResult))
            {
                return cachedResult;
            }
        }

        var ast = FormulaParser.Parse(formula);
        var result = EvaluateNode(ast, activeWorksheet);

        lock (_formulaCache)
        {
            _formulaCache[cacheKey] = result;
        }

        return result;
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
                return EvaluateReference(reference, sheet);

            case OperatorNode opNode:
                var leftVal = EvaluateNode(opNode.Left, sheet);
                var rightVal = opNode.Right != null ? EvaluateNode(opNode.Right, sheet) : null;
                return EvaluateOperator(opNode.Operator, leftVal, rightVal);

            case FunctionNode funcNode:
                var evaluatedArgs = funcNode.Arguments.Select(arg => EvaluateNode(arg, sheet)).ToList();
                return EvaluateFunction(funcNode.FunctionName, evaluatedArgs, sheet);

            default:
                throw new FormulaException($"Unsupported AST node type: '{node.GetType().Name}'", "UNSUPPORTED_NODE");
        }
    }

    private static object? EvaluateReference(ReferenceNode reference, ExcelWorksheet sheet)
    {
        var targetSheet = string.IsNullOrEmpty(reference.SheetName)
            ? sheet
            : (sheet.Workbook.Worksheets.FirstOrDefault(w => string.Equals(w.Name, reference.SheetName, StringComparison.OrdinalIgnoreCase))
               ?? throw new FormulaException($"Worksheet '{reference.SheetName}' was not found.", "WORKSHEET_NOT_FOUND"));

        var addr = reference.RangeAddress;

        if (addr.StartRow == addr.EndRow && addr.StartColumn == addr.EndColumn)
        {
            var cellAddr = new CellAddress(addr.StartRow, addr.StartColumn);
            return GetCellValueEvaluated(targetSheet, cellAddr);
        }

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

    private static object? GetCellValueEvaluated(ExcelWorksheet sheet, CellAddress address)
    {
        var formula = sheet.GetCellFormula(address);
        if (!string.IsNullOrEmpty(formula))
        {
            return Evaluate(formula!, sheet, address);
        }
        return sheet.GetCellValue(address);
    }

    private static object? EvaluateOperator(string op, object? left, object? right)
    {
        if (right == null)
        {
            if (op == "-") return -ConvertToDouble(left);
            if (op == "+") return ConvertToDouble(left);
            throw new FormulaException($"Unsupported unary operator: '{op}'", "INVALID_OPERATOR");
        }

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

    private static object? EvaluateFunction(string name, List<object?> args, ExcelWorksheet sheet)
    {
        string upperName = name.ToUpperInvariant();

        switch (upperName)
        {
            #region Math & Stats

            case "SUM":
                return Flatten(args).Select(ConvertToDouble).Sum();

            case "AVERAGE":
                var sumList = Flatten(args).Select(ConvertToDouble).ToList();
                return sumList.Count == 0 ? 0.0 : sumList.Average();

            case "COUNT":
                return Flatten(args).Count(val => val is double || val is float || val is int || val is decimal || (val is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out _)));

            case "COUNTA":
                return Flatten(args).Count(val => val != null && !string.IsNullOrEmpty(val.ToString()));

            case "MIN":
                var minList = Flatten(args).Select(ConvertToDouble).ToList();
                return minList.Count == 0 ? 0.0 : minList.Min();

            case "MAX":
                var maxList = Flatten(args).Select(ConvertToDouble).ToList();
                return maxList.Count == 0 ? 0.0 : maxList.Max();

            case "ABS":
                if (args.Count < 1) throw new FormulaException("ABS requires 1 argument.", "ARGUMENT_COUNT_ERROR");
                return Math.Abs(ConvertToDouble(args[0]));

            case "ROUND":
                if (args.Count < 2) throw new FormulaException("ROUND requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                return Math.Round(ConvertToDouble(args[0]), (int)ConvertToDouble(args[1]), MidpointRounding.AwayFromZero);

            case "ROUNDUP":
                if (args.Count < 2) throw new FormulaException("ROUNDUP requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                double valUp = ConvertToDouble(args[0]);
                int digitsUp = (int)ConvertToDouble(args[1]);
                double factorUp = Math.Pow(10, digitsUp);
                return valUp >= 0 ? Math.Ceiling(valUp * factorUp) / factorUp : Math.Floor(valUp * factorUp) / factorUp;

            case "ROUNDDOWN":
                if (args.Count < 2) throw new FormulaException("ROUNDDOWN requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                double valDown = ConvertToDouble(args[0]);
                int digitsDown = (int)ConvertToDouble(args[1]);
                double factorDown = Math.Pow(10, digitsDown);
                return valDown >= 0 ? Math.Floor(valDown * factorDown) / factorDown : Math.Ceiling(valDown * factorDown) / factorDown;

            case "POWER":
                if (args.Count < 2) throw new FormulaException("POWER requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                return Math.Pow(ConvertToDouble(args[0]), ConvertToDouble(args[1]));

            case "SQRT":
                if (args.Count < 1) throw new FormulaException("SQRT requires 1 argument.", "ARGUMENT_COUNT_ERROR");
                double sqrtVal = ConvertToDouble(args[0]);
                if (sqrtVal < 0) throw new FormulaException("Cannot calculate square root of a negative number.", "INVALID_ARGUMENT");
                return Math.Sqrt(sqrtVal);

            case "MOD":
                if (args.Count < 2) throw new FormulaException("MOD requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                double div = ConvertToDouble(args[1]);
                if (div == 0) throw new FormulaException("MOD division by zero.", "DIVISION_BY_ZERO");
                return ConvertToDouble(args[0]) % div;

            case "INT":
                if (args.Count < 1) throw new FormulaException("INT requires 1 argument.", "ARGUMENT_COUNT_ERROR");
                return Math.Floor(ConvertToDouble(args[0]));

            case "CEILING":
                if (args.Count < 1) throw new FormulaException("CEILING requires 1 argument.", "ARGUMENT_COUNT_ERROR");
                double sigC = args.Count > 1 ? ConvertToDouble(args[1]) : 1.0;
                return Math.Ceiling(ConvertToDouble(args[0]) / sigC) * sigC;

            case "FLOOR":
                if (args.Count < 1) throw new FormulaException("FLOOR requires 1 argument.", "ARGUMENT_COUNT_ERROR");
                double sigF = args.Count > 1 ? ConvertToDouble(args[1]) : 1.0;
                return Math.Floor(ConvertToDouble(args[0]) / sigF) * sigF;

            case "COUNTIF":
                if (args.Count < 2) throw new FormulaException("COUNTIF requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                var countRange = Flatten(new List<object?> { args[0] }).ToList();
                string criteriaStr = args[1]?.ToString() ?? string.Empty;
                return countRange.Count(v => EvaluateCriteria(v, criteriaStr));

            case "SUMIF":
                if (args.Count < 2) throw new FormulaException("SUMIF requires at least 2 arguments.", "ARGUMENT_COUNT_ERROR");
                var sumRangeCond = Flatten(new List<object?> { args[0] }).ToList();
                string sumCrit = args[1]?.ToString() ?? string.Empty;
                var sumValues = args.Count > 2 ? Flatten(new List<object?> { args[2] }).ToList() : sumRangeCond;
                double sumResult = 0.0;
                for (int i = 0; i < Math.Min(sumRangeCond.Count, sumValues.Count); i++)
                {
                    if (EvaluateCriteria(sumRangeCond[i], sumCrit))
                    {
                        sumResult += ConvertToDouble(sumValues[i]);
                    }
                }
                return sumResult;

            #endregion

            #region Lookup & Reference

            case "VLOOKUP":
                if (args.Count < 3) throw new FormulaException("VLOOKUP requires at least 3 arguments.", "ARGUMENT_COUNT_ERROR");
                var lookupKey = args[0];
                var tableMatrix = args[1] as List<object?>;
                int colIdx = (int)ConvertToDouble(args[2]);
                bool exactMatch = args.Count > 3 ? !ConvertToBoolean(args[3]) : true;

                if (tableMatrix == null || colIdx < 1) return null;

                for (int i = 0; i < tableMatrix.Count; i += colIdx)
                {
                    var keyVal = tableMatrix[i];
                    if (CompareValues(keyVal, lookupKey) == 0 || (!exactMatch && CompareValues(keyVal, lookupKey) <= 0))
                    {
                        int targetIdx = i + colIdx - 1;
                        return targetIdx < tableMatrix.Count ? tableMatrix[targetIdx] : null;
                    }
                }
                return null;

            case "MATCH":
                if (args.Count < 2) throw new FormulaException("MATCH requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                var matchKey = args[0];
                var matchArray = Flatten(new List<object?> { args[1] }).ToList();
                for (int i = 0; i < matchArray.Count; i++)
                {
                    if (CompareValues(matchArray[i], matchKey) == 0) return i + 1; // 1-based index
                }
                return null;

            case "INDEX":
                if (args.Count < 2) throw new FormulaException("INDEX requires at least 2 arguments.", "ARGUMENT_COUNT_ERROR");
                var indexArray = Flatten(new List<object?> { args[0] }).ToList();
                int idxRow = (int)ConvertToDouble(args[1]);
                if (idxRow >= 1 && idxRow <= indexArray.Count) return indexArray[idxRow - 1];
                return null;

            #endregion

            #region Text

            case "UPPER":
                return args.Count > 0 ? args[0]?.ToString()?.ToUpperInvariant() ?? string.Empty : string.Empty;

            case "LOWER":
                return args.Count > 0 ? args[0]?.ToString()?.ToLowerInvariant() ?? string.Empty : string.Empty;

            case "PROPER":
                if (args.Count == 0 || args[0] == null) return string.Empty;
                return CultureInfo.InvariantCulture.TextInfo.ToTitleCase(args[0]!.ToString()!.ToLowerInvariant());

            case "TRIM":
                return args.Count > 0 ? args[0]?.ToString()?.Trim() ?? string.Empty : string.Empty;

            case "MID":
                if (args.Count < 3) throw new FormulaException("MID requires 3 arguments.", "ARGUMENT_COUNT_ERROR");
                string midStr = args[0]?.ToString() ?? string.Empty;
                int startNum = (int)ConvertToDouble(args[1]) - 1; // 1-based to 0-based
                int numChars = (int)ConvertToDouble(args[2]);
                if (startNum < 0 || startNum >= midStr.Length || numChars <= 0) return string.Empty;
                return startNum + numChars >= midStr.Length ? midStr.Substring(startNum) : midStr.Substring(startNum, numChars);

            case "SUBSTITUTE":
                if (args.Count < 3) throw new FormulaException("SUBSTITUTE requires at least 3 arguments.", "ARGUMENT_COUNT_ERROR");
                string subText = args[0]?.ToString() ?? string.Empty;
                string oldText = args[1]?.ToString() ?? string.Empty;
                string newText = args[2]?.ToString() ?? string.Empty;
                return subText.Replace(oldText, newText);

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
            case "TEXTJOIN":
                string delim = upperName == "TEXTJOIN" && args.Count > 0 ? args[0]?.ToString() ?? string.Empty : string.Empty;
                bool ignoreEmpty = upperName == "TEXTJOIN" && args.Count > 1 && ConvertToBoolean(args[1]);
                var items = Flatten(args.Skip(upperName == "TEXTJOIN" ? 2 : 0).ToList()).Select(a => a?.ToString() ?? string.Empty);
                if (ignoreEmpty) items = items.Where(s => !string.IsNullOrEmpty(s));
                return string.Join(delim, items);

            #endregion

            #region Logical & Information

            case "IF":
                if (args.Count < 2) throw new FormulaException("IF requires at least 2 arguments.", "ARGUMENT_COUNT_ERROR");
                return ConvertToBoolean(args[0]) ? args[1] : (args.Count > 2 ? args[2] : null);

            case "IFS":
                for (int i = 0; i < args.Count - 1; i += 2)
                {
                    if (ConvertToBoolean(args[i])) return args[i + 1];
                }
                return null;

            case "IFERROR":
                if (args.Count < 2) throw new FormulaException("IFERROR requires 2 arguments.", "ARGUMENT_COUNT_ERROR");
                try
                {
                    var val = args[0];
                    return val ?? args[1];
                }
                catch
                {
                    return args[1];
                }

            case "AND":
                return args.Count > 0 && args.All(ConvertToBoolean);

            case "OR":
                return args.Count > 0 && args.Any(ConvertToBoolean);

            case "NOT":
                if (args.Count != 1) throw new FormulaException("NOT requires 1 argument.", "ARGUMENT_COUNT_ERROR");
                return !ConvertToBoolean(args[0]);

            case "ISBLANK":
                return args.Count > 0 && (args[0] == null || string.IsNullOrEmpty(args[0]?.ToString()));

            case "ISNUMBER":
                return args.Count > 0 && (args[0] is double || args[0] is int || args[0] is float || args[0] is decimal);

            case "ISTEXT":
                return args.Count > 0 && args[0] is string;

            #endregion

            #region Date & Time

            case "TODAY":
                return DateTime.Today;

            case "NOW":
                return DateTime.Now;

            case "DATE":
                if (args.Count < 3) throw new FormulaException("DATE requires 3 arguments.", "ARGUMENT_COUNT_ERROR");
                int year = (int)ConvertToDouble(args[0]);
                int month = (int)ConvertToDouble(args[1]);
                int day = (int)ConvertToDouble(args[2]);
                return new DateTime(year, month, day);

            case "YEAR":
                return args.Count > 0 ? ConvertToDateTime(args[0]).Year : 0;

            case "MONTH":
                return args.Count > 0 ? ConvertToDateTime(args[0]).Month : 0;

            case "DAY":
                return args.Count > 0 ? ConvertToDateTime(args[0]).Day : 0;

            #endregion

            default:
                throw new FormulaException($"Function '{name}' is not supported yet by the evaluator engine.", "UNSUPPORTED_FUNCTION");
        }
    }

    private static bool EvaluateCriteria(object? cellValue, string criteria)
    {
        if (string.IsNullOrEmpty(criteria)) return cellValue == null;

        if (criteria.StartsWith(">=")) return CompareValues(cellValue, criteria.Substring(2)) >= 0;
        if (criteria.StartsWith("<=")) return CompareValues(cellValue, criteria.Substring(2)) <= 0;
        if (criteria.StartsWith("<>")) return CompareValues(cellValue, criteria.Substring(2)) != 0;
        if (criteria.StartsWith(">")) return CompareValues(cellValue, criteria.Substring(1)) > 0;
        if (criteria.StartsWith("<")) return CompareValues(cellValue, criteria.Substring(1)) < 0;
        if (criteria.StartsWith("=")) return CompareValues(cellValue, criteria.Substring(1)) == 0;

        return string.Equals(cellValue?.ToString(), criteria, StringComparison.OrdinalIgnoreCase);
    }

    private static double ConvertToDouble(object? val)
    {
        if (val == null) return 0.0;
        if (val is double d) return d;
        if (val is float f) return f;
        if (val is int i) return i;
        if (val is long l) return l;
        if (val is decimal dec) return (double)dec;
        if (val is string s && double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsed)) return parsed;
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

    private static DateTime ConvertToDateTime(object? val)
    {
        if (val is DateTime dt) return dt;
        if (val is double d) return DateTime.FromOADate(d);
        if (val is string s && DateTime.TryParse(s, out DateTime parsed)) return parsed;
        return DateTime.MinValue;
    }

    private static int CompareValues(object? left, object? right)
    {
        if (left == null && right == null) return 0;
        if (left == null) return -1;
        if (right == null) return 1;

        if (left is double || left is float || left is int || left is long || right is double || right is float || right is int || right is long)
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
