using FluentAssertions;
using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;
using SmartExcelKit.Formula;
using Xunit;

namespace SmartExcelKit.Tests;

public class FormulaEngineTests
{
    [Fact]
    public void Tokenizer_ShouldParseTokens()
    {
        // Arrange
        string formula = "=SUM(A1:B2) + 12.34";

        // Act
        var tokens = FormulaTokenizer.Tokenize(formula);

        // Assert
        tokens.Count.Should().Be(6);
        tokens[0].Type.Should().Be(FormulaTokenType.Identifier);
        tokens[0].Value.Should().Be("SUM");
        tokens[1].Type.Should().Be(FormulaTokenType.LeftParenthesis);
        tokens[2].Type.Should().Be(FormulaTokenType.Reference);
        tokens[2].Value.Should().Be("A1:B2");
        tokens[3].Type.Should().Be(FormulaTokenType.RightParenthesis);
        tokens[4].Type.Should().Be(FormulaTokenType.Operator);
        tokens[4].Value.Should().Be("+");
        tokens[5].Type.Should().Be(FormulaTokenType.Number);
        tokens[5].Value.Should().Be("12.34");
    }

    [Fact]
    public void Parser_ShouldCreateCorrectAST()
    {
        // Arrange
        string formula = "=10 + 5 * 2";

        // Act
        var ast = FormulaParser.Parse(formula);

        // Assert
        ast.Should().BeOfType<OperatorNode>();
        var opNode = (OperatorNode)ast;
        opNode.Operator.Should().Be("+");
        opNode.Left.Should().BeOfType<LiteralNode>();
        opNode.Right.Should().BeOfType<OperatorNode>();

        var rightOp = (OperatorNode)opNode.Right!;
        rightOp.Operator.Should().Be("*");
    }

    [Fact]
    public void Evaluator_ShouldEvaluateBasicArithmetic()
    {
        // Arrange
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");

        // Act
        var result = FormulaEvaluator.Evaluate("=10 + 5 * 2", sheet, new CellAddress(1, 1));

        // Assert
        result.Should().Be(20.0);
    }

    [Fact]
    public void Evaluator_ShouldEvaluateFunctions()
    {
        // Arrange
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cell("A1").Value = 10.0;
        sheet.Cell("A2").Value = 20.0;
        sheet.Cell("A3").Value = 30.0;

        // Act
        var sumResult = FormulaEvaluator.Evaluate("=SUM(A1:A3)", sheet, new CellAddress(4, 1));
        var avgResult = FormulaEvaluator.Evaluate("=AVERAGE(A1:A3)", sheet, new CellAddress(5, 1));
        var minResult = FormulaEvaluator.Evaluate("=MIN(A1:A3)", sheet, new CellAddress(6, 1));

        // Assert
        sumResult.Should().Be(60.0);
        avgResult.Should().Be(20.0);
        minResult.Should().Be(10.0);
    }

    [Fact]
    public void Evaluator_ShouldEvaluateLogicalIf()
    {
        // Arrange
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        sheet.Cell("A1").Value = 5.0;

        // Act & Assert
        FormulaEvaluator.Evaluate("=IF(A1 > 3, \"Yes\", \"No\")", sheet, new CellAddress(2, 1)).Should().Be("Yes");
        FormulaEvaluator.Evaluate("=IF(A1 < 2, \"Yes\", \"No\")", sheet, new CellAddress(2, 1)).Should().Be("No");
    }

    [Fact]
    public void DependencyGraph_ShouldDetectCycleAndThrow()
    {
        // Arrange
        var graph = new DependencyGraph();

        // Act
        graph.SetDependencies("Sheet1!A1", ["Sheet1!B1"]);
        graph.SetDependencies("Sheet1!B1", ["Sheet1!C1"]);

        // Act & Assert
        var act = () => graph.SetDependencies("Sheet1!C1", ["Sheet1!A1"]); // Cycle: A1 -> B1 -> C1 -> A1
        act.Should().Throw<FormulaException>().WithMessage("*Circular reference*");
    }
}
