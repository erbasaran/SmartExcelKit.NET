namespace SmartExcelKit.Formula;

/// <summary>
/// Tokenizer that splits a formula expression into structured tokens.
/// </summary>
internal static class FormulaTokenizer
{
    private static readonly FormulaToken LeftParenthesisToken = new(FormulaTokenType.LeftParenthesis, "(");
    private static readonly FormulaToken RightParenthesisToken = new(FormulaTokenType.RightParenthesis, ")");
    private static readonly FormulaToken CommaToken = new(FormulaTokenType.Comma, ",");

    private static readonly FormulaToken PlusToken = new(FormulaTokenType.Operator, "+");
    private static readonly FormulaToken MinusToken = new(FormulaTokenType.Operator, "-");
    private static readonly FormulaToken MultiplyToken = new(FormulaTokenType.Operator, "*");
    private static readonly FormulaToken DivideToken = new(FormulaTokenType.Operator, "/");
    private static readonly FormulaToken EqualToken = new(FormulaTokenType.Operator, "=");
    private static readonly FormulaToken PowerToken = new(FormulaTokenType.Operator, "^");

    /// <summary>
    /// Tokenizes a formula string.
    /// </summary>
    /// <param name="formula">The formula string (e.g. "=SUM(A1:B2) + 10").</param>
    /// <returns>A list of token results.</returns>
    public static List<FormulaToken> Tokenize(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
            return [];

        var tokens = new List<FormulaToken>();
        int i = 0;

        // Skip leading '=' if present
        if (formula[0] == '=')
        {
            i++;
        }

        while (i < formula.Length)
        {
            char c = formula[i];

            if (char.IsWhiteSpace(c))
            {
                i++;
                continue;
            }

            // Parentheses, Comma
            if (c == '(')
            {
                tokens.Add(LeftParenthesisToken);
                i++;
                continue;
            }
            if (c == ')')
            {
                tokens.Add(RightParenthesisToken);
                i++;
                continue;
            }
            if (c == ',')
            {
                tokens.Add(CommaToken);
                i++;
                continue;
            }

            // Operators (including multi-character operators like <=, >=, <>)
            if (IsOperatorChar(c))
            {
                FormulaToken? staticToken = c switch
                {
                    '+' => PlusToken,
                    '-' => MinusToken,
                    '*' => MultiplyToken,
                    '/' => DivideToken,
                    '^' => PowerToken,
                    '=' => EqualToken,
                    _ => null
                };

                if (staticToken != null && (i + 1 >= formula.Length || !IsOperatorChar(formula[i + 1])))
                {
                    tokens.Add(staticToken);
                    i++;
                    continue;
                }

                // Check for multi-character operators
                string opStr;
                if (i + 1 < formula.Length)
                {
                    char next = formula[i + 1];
                    if ((c == '<' && (next == '=' || next == '>')) || (c == '>' && next == '='))
                    {
                        opStr = formula.Substring(i, 2);
                        i += 2;
                    }
                    else
                    {
                        opStr = c.ToString();
                        i++;
                    }
                }
                else
                {
                    opStr = c.ToString();
                    i++;
                }

                tokens.Add(new FormulaToken(FormulaTokenType.Operator, opStr));
                continue;
            }

            // String literals: "hello"
            if (c == '"')
            {
                int start = i + 1; // skip opening quote
                i++;

                // Fast path: check if it's a simple string without escaped quotes
                bool hasEscapedQuotes = false;
                while (i < formula.Length && formula[i] != '"')
                {
                    i++;
                }

                if (i < formula.Length && formula[i] == '"')
                {
                    // Check if there is an escaped quote right after it
                    if (i + 1 < formula.Length && formula[i + 1] == '"')
                    {
                        hasEscapedQuotes = true;
                    }
                }

                if (!hasEscapedQuotes)
                {
                    // No escaped quotes, extract substring directly
                    string strVal = i > start ? formula.Substring(start, i - start) : string.Empty;
                    if (i < formula.Length) i++; // skip closing quote
                    tokens.Add(new FormulaToken(FormulaTokenType.String, strVal));
                }
                else
                {
                    // Slow path: handle escaped quotes
                    i = start - 1; // reset
                    var strSb = new System.Text.StringBuilder();
                    i++; // skip opening quote
                    while (i < formula.Length)
                    {
                        if (formula[i] == '"')
                        {
                            if (i + 1 < formula.Length && formula[i + 1] == '"')
                            {
                                strSb.Append('"');
                                i += 2;
                            }
                            else
                            {
                                i++; // skip closing quote
                                break;
                            }
                        }
                        else
                        {
                            strSb.Append(formula[i]);
                            i++;
                        }
                    }
                    tokens.Add(new FormulaToken(FormulaTokenType.String, strSb.ToString()));
                }
                continue;
            }

            // Numbers
            if (char.IsDigit(c) || c == '.')
            {
                int start = i;
                bool hasDot = false;
                while (i < formula.Length && (char.IsDigit(formula[i]) || (!hasDot && formula[i] == '.')))
                {
                    if (formula[i] == '.') hasDot = true;
                    i++;
                }
                string numStr = formula.Substring(start, i - start);
                tokens.Add(new FormulaToken(FormulaTokenType.Number, numStr));
                continue;
            }

            // Identifiers / Cell References
            if (char.IsLetter(c) || c == '$' || c == '_')
            {
                int start = i;
                while (i < formula.Length && (char.IsLetterOrDigit(formula[i]) || formula[i] == '$' || formula[i] == '_' || formula[i] == ':'))
                {
                    i++;
                }

                string idStr = formula.Substring(start, i - start);

                if (string.Equals(idStr, "TRUE", StringComparison.OrdinalIgnoreCase) || string.Equals(idStr, "FALSE", StringComparison.OrdinalIgnoreCase))
                {
                    tokens.Add(new FormulaToken(FormulaTokenType.Boolean, idStr.ToUpperInvariant()));
                }
                else if (IsCellReference(idStr))
                {
                    tokens.Add(new FormulaToken(FormulaTokenType.Reference, idStr.ToUpperInvariant()));
                }
                else
                {
                    tokens.Add(new FormulaToken(FormulaTokenType.Identifier, idStr.ToUpperInvariant()));
                }
                continue;
            }

            // Catch-all: unexpected character
            tokens.Add(new FormulaToken(FormulaTokenType.Error, c.ToString()));
            i++;
        }

        return tokens;
    }

    private static bool IsOperatorChar(char c)
    {
        return c == '+' || c == '-' || c == '*' || c == '/' || c == '^' || c == '=' || c == '<' || c == '>' || c == '&' || c == '!';
    }

    private static bool IsCellReference(string s)
    {
        int i = 0;
        if (i < s.Length && s[i] == '$') i++;
        if (i >= s.Length || !char.IsLetter(s[i])) return false;

        // Skip letters
        while (i < s.Length && char.IsLetter(s[i])) i++;
        if (i < s.Length && s[i] == '$') i++;
        // Must contain digits
        if (i >= s.Length || !char.IsDigit(s[i])) return false;
        while (i < s.Length && char.IsDigit(s[i])) i++;

        if (i == s.Length) return true; // Single cell reference

        // Range separator
        if (i < s.Length && s[i] == ':')
        {
            i++;
            if (i < s.Length && s[i] == '$') i++;
            if (i >= s.Length || !char.IsLetter(s[i])) return false;
            while (i < s.Length && char.IsLetter(s[i])) i++;
            if (i < s.Length && s[i] == '$') i++;
            if (i >= s.Length || !char.IsDigit(s[i])) return false;
            while (i < s.Length && char.IsDigit(s[i])) i++;
            return i == s.Length;
        }

        return false;
    }
}
