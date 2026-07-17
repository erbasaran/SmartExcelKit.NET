using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;

namespace SmartExcelKit.Formula;

/// <summary>
/// Parser that translates tokens into a Formula AST.
/// </summary>
internal sealed class FormulaParser
{
    private readonly List<FormulaToken> _tokens;
    private int _index;

    private FormulaParser(List<FormulaToken> tokens)
    {
        _tokens = tokens;
        _index = 0;
    }

    /// <summary>
    /// Parses a formula expression string into an AST Node.
    /// </summary>
    public static FormulaNode Parse(string formula)
    {
        var tokens = FormulaTokenizer.Tokenize(formula);
        if (tokens.Count == 0)
        {
            throw new FormulaException("Empty formula string.", "EMPTY_FORMULA");
        }
        var parser = new FormulaParser(tokens);
        var node = parser.ParseExpression();
        if (parser._index < tokens.Count)
        {
            throw new FormulaException($"Unexpected token at end of formula: '{tokens[parser._index].Value}'", "UNEXPECTED_TOKEN");
        }
        return node;
    }

    private FormulaToken Peek()
    {
        if (_index >= _tokens.Count) return new FormulaToken(FormulaTokenType.Error, "EOF");
        return _tokens[_index];
    }

    private FormulaToken Consume()
    {
        var token = Peek();
        _index++;
        return token;
    }

    private void Expect(FormulaTokenType type, string? val = null)
    {
        var token = Peek();
        if (token.Type != type || (val != null && !string.Equals(token.Value, val, StringComparison.OrdinalIgnoreCase)))
        {
            throw new FormulaException($"Expected token of type '{type}' with value '{val ?? "any"}', but got '{token.Type}' ('{token.Value}').", "SYNTAX_ERROR");
        }
        _index++;
    }

    // Expression -> Comparison
    private FormulaNode ParseExpression()
    {
        return ParseComparison();
    }

    // Comparison -> Concat ( ( "=" | "<>" | "<" | ">" | "<=" | ">=" ) Concat )*
    private FormulaNode ParseComparison()
    {
        var node = ParseConcat();
        while (true)
        {
            var token = Peek();
            if (token.Type == FormulaTokenType.Operator &&
                (token.Value == "=" || token.Value == "<>" || token.Value == "<" || token.Value == ">" || token.Value == "<=" || token.Value == ">="))
            {
                Consume();
                var right = ParseConcat();
                node = new OperatorNode(token.Value, node, right);
            }
            else
            {
                break;
            }
        }
        return node;
    }

    // Concat -> Additive ( "&" Additive )*
    private FormulaNode ParseConcat()
    {
        var node = ParseAdditive();
        while (true)
        {
            var token = Peek();
            if (token.Type == FormulaTokenType.Operator && token.Value == "&")
            {
                Consume();
                var right = ParseAdditive();
                node = new OperatorNode("&", node, right);
            }
            else
            {
                break;
            }
        }
        return node;
    }

    // Additive -> Multiplicative ( ( "+" | "-" ) Multiplicative )*
    private FormulaNode ParseAdditive()
    {
        var node = ParseMultiplicative();
        while (true)
        {
            var token = Peek();
            if (token.Type == FormulaTokenType.Operator && (token.Value == "+" || token.Value == "-"))
            {
                Consume();
                var right = ParseMultiplicative();
                node = new OperatorNode(token.Value, node, right);
            }
            else
            {
                break;
            }
        }
        return node;
    }

    // Multiplicative -> Exponent ( ( "*" | "/" ) Exponent )*
    private FormulaNode ParseMultiplicative()
    {
        var node = ParseExponent();
        while (true)
        {
            var token = Peek();
            if (token.Type == FormulaTokenType.Operator && (token.Value == "*" || token.Value == "/"))
            {
                Consume();
                var right = ParseExponent();
                node = new OperatorNode(token.Value, node, right);
            }
            else
            {
                break;
            }
        }
        return node;
    }

    // Exponent -> Unary ( "^" Unary )*
    private FormulaNode ParseExponent()
    {
        var node = ParseUnary();
        while (true)
        {
            var token = Peek();
            if (token.Type == FormulaTokenType.Operator && token.Value == "^")
            {
                Consume();
                var right = ParseUnary();
                node = new OperatorNode("^", node, right);
            }
            else
            {
                break;
            }
        }
        return node;
    }

    // Unary -> ( "+" | "-" )? Primary
    private FormulaNode ParseUnary()
    {
        var token = Peek();
        if (token.Type == FormulaTokenType.Operator && (token.Value == "+" || token.Value == "-"))
        {
            Consume();
            var primary = ParsePrimary();
            return new OperatorNode(token.Value, primary); // Unary operator representation
        }
        return ParsePrimary();
    }

    // Primary -> Literal | Reference | FunctionCall | "(" Expression ")"
    private FormulaNode ParsePrimary()
    {
        var token = Peek();

        if (token.Type == FormulaTokenType.Number)
        {
            Consume();
            if (double.TryParse(token.Value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double d))
                return new LiteralNode(d);
            return new LiteralNode(token.Value);
        }

        if (token.Type == FormulaTokenType.String)
        {
            Consume();
            return new LiteralNode(token.Value);
        }

        if (token.Type == FormulaTokenType.Boolean)
        {
            Consume();
            return new LiteralNode(string.Equals(token.Value, "TRUE", StringComparison.OrdinalIgnoreCase));
        }

        if (token.Type == FormulaTokenType.LeftParenthesis)
        {
            Consume();
            var node = ParseExpression();
            Expect(FormulaTokenType.RightParenthesis, ")");
            return node;
        }

        // Cross-sheet reference support check: Identifier("Sheet1") + Operator("!") + Reference("A1:C3")
        if (token.Type == FormulaTokenType.Identifier)
        {
            if (_index + 2 < _tokens.Count && _tokens[_index + 1].Type == FormulaTokenType.Operator && _tokens[_index + 1].Value == "!" && _tokens[_index + 2].Type == FormulaTokenType.Reference)
            {
                string sheetName = token.Value;
                Consume(); // Consume Identifier
                Consume(); // Consume '!'
                var refToken = Consume(); // Consume Reference
                var rangeAddr = ExcelRangeAddress.Parse(refToken.Value);
                return new ReferenceNode(rangeAddr, sheetName);
            }

            // Normal Function Call: Identifier("SUM") + LeftParenthesis
            if (_index + 1 < _tokens.Count && _tokens[_index + 1].Type == FormulaTokenType.LeftParenthesis)
            {
                string funcName = token.Value;
                Consume(); // Consume Identifier
                Consume(); // Consume '('

                var args = new List<FormulaNode>();
                if (Peek().Type != FormulaTokenType.RightParenthesis)
                {
                    args.Add(ParseExpression());
                    while (Peek().Type == FormulaTokenType.Comma)
                    {
                        Consume(); // Consume ','
                        args.Add(ParseExpression());
                    }
                }
                Expect(FormulaTokenType.RightParenthesis, ")");
                return new FunctionNode(funcName, args);
            }
        }

        if (token.Type == FormulaTokenType.Reference)
        {
            Consume();
            var rangeAddr = ExcelRangeAddress.Parse(token.Value);
            return new ReferenceNode(rangeAddr);
        }

        throw new FormulaException($"Invalid or unexpected token in formula: '{token.Value}'", "PARSING_ERROR");
    }
}
