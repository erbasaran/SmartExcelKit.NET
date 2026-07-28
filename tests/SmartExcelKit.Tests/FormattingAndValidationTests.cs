using FluentAssertions;
using SmartExcelKit.Formatting;
using SmartExcelKit.Styles;
using SmartExcelKit.Validation;
using Xunit;

namespace SmartExcelKit.Tests;

public class FormattingAndValidationTests
{
    [Fact]
    public void ConditionalFormatting_AddRules_StoresCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("FormattingSheet");

        var rule1 = ws.ConditionalFormatting.AddCellValueRule("A1:A10", ConditionalFormattingOperator.GreaterThan, "100", new ExcelStyle(font: new ExcelFont(bold: true, color: "FF0000")));
        var rule2 = ws.ConditionalFormatting.AddDataBar("B1:B10", "00FF00");

        ws.ConditionalFormatting.Count.Should().Be(2);
        rule1.RuleType.Should().Be(ConditionalFormattingType.CellValue);
        rule2.RuleType.Should().Be(ConditionalFormattingType.DataBar);
    }

    [Fact]
    public void DataValidation_AddValidations_StoresCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("ValidationSheet");

        var val1 = ws.DataValidations.AddListValidation("A1:A10", new[] { "Option 1", "Option 2", "Option 3" });
        var val2 = ws.DataValidations.AddWholeNumberValidation("B1:B10", ValidationOperator.Between, 1, 100);

        ws.DataValidations.Count.Should().Be(2);
        val1.ValidationType.Should().Be(ValidationType.List);
        val2.ValidationType.Should().Be(ValidationType.WholeNumber);
    }
}
