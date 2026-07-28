using FluentAssertions;
using SmartExcelKit.Core;
using SmartExcelKit.Drawings;
using Xunit;

namespace SmartExcelKit.Tests;

public class RichContentAndDrawingsTests
{
    [Fact]
    public void RichText_Comments_Hyperlinks_WorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("RichSheet");

        var rich = new RichText();
        rich.AddBold("Hello ");
        rich.AddItalic("World");
        ws[1, 1].RichText = rich;

        ws[1, 1].GetString().Should().Be("Hello World");
        ws[1, 1].RichText.Should().NotBeNull();

        ws[1, 2].CommentObject = new ExcelComment("Test comment", "Author");
        ws[1, 2].CommentObject!.Text.Should().Be("Test comment");

        ws[1, 3].HyperlinkObject = ExcelHyperlink.Internal("RichSheet", "A1");
        ws[1, 3].HyperlinkObject!.HyperlinkType.Should().Be(HyperlinkType.InternalReference);
    }

    [Fact]
    public void Drawings_Charts_Images_PivotTables_CanBeAdded()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("DrawingsSheet");

        var chart = new ExcelChart(ChartType.Column, 5, 2);
        chart.Title = "Sales Overview";
        chart.AddSeries("Revenue", "RichSheet!B1:B10");
        ws.Charts.Add(chart);

        ws.Charts.Count.Should().Be(1);
        ws.Charts[0].Title.Should().Be("Sales Overview");

        var pivot = new ExcelPivotTable("SalesPivot", "Data!A1:C100", new CellAddress(10, 1));
        pivot.AddRowField("Region");
        pivot.AddDataField("Sales", PivotSummaryFunction.Sum);
        ws.PivotTables.Add(pivot);

        ws.PivotTables.Count.Should().Be(1);
        ws.PivotTables[0].Name.Should().Be("SalesPivot");
    }
}
