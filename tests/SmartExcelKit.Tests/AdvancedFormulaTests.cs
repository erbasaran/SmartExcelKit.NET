using FluentAssertions;
using Xunit;

namespace SmartExcelKit.Tests;

public class AdvancedFormulaTests
{
    [Fact]
    public void Evaluate_MathFunctions_ReturnsCorrectValues()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("MathSheet");

        ws[1, 1].Value = -15.5;
        ws[1, 2].Formula = "=ABS(A1)";
        ws[1, 3].Formula = "=ROUND(A1, 0)";
        ws[1, 4].Formula = "=POWER(2, 3)";
        ws[1, 5].Formula = "=SQRT(16)";

        workbook.Recalculate();

        ws[1, 2].Value.Should().Be(15.5);
        ws[1, 3].Value.Should().Be(-16.0);
        ws[1, 4].Value.Should().Be(8.0);
        ws[1, 5].Value.Should().Be(4.0);
    }

    [Fact]
    public void Evaluate_TextFunctions_ReturnsCorrectStrings()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("TextSheet");

        ws[1, 1].Value = "  hello world  ";
        ws[1, 2].Formula = "=TRIM(A1)";
        ws[1, 3].Formula = "=UPPER(TRIM(A1))";
        ws[1, 4].Formula = "=LEN(UPPER(TRIM(A1)))";
        ws[1, 5].Formula = "=PROPER(\"john doe\")";

        workbook.Recalculate();

        ws[1, 2].Value.Should().Be("hello world");
        ws[1, 3].Value.Should().Be("HELLO WORLD");
        ws[1, 4].Value.Should().Be(11);
        ws[1, 5].Value.Should().Be("John Doe");
    }

    [Fact]
    public void Evaluate_LookupAndReference_ReturnsMatchingCell()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("LookupSheet");

        ws[1, 1].Value = 101; ws[1, 2].Value = "Alice";
        ws[2, 1].Value = 102; ws[2, 2].Value = "Bob";
        ws[3, 1].Value = 103; ws[3, 2].Value = "Charlie";

        ws[4, 1].Formula = "=MATCH(102, A1:A3)";

        workbook.Recalculate();

        ws[4, 1].Value.Should().Be(2);
    }
}
