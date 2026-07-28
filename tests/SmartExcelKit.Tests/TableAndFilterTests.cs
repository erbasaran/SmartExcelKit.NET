using FluentAssertions;
using SmartExcelKit.Tables;
using Xunit;

namespace SmartExcelKit.Tests;

public class TableAndFilterTests
{
    [Fact]
    public void Tables_AddAndManipulate_WorksAsExpected()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Data");

        ws[1, 1].Value = "ID";
        ws[1, 2].Value = "Amount";
        ws[2, 1].Value = 1;
        ws[2, 2].Value = 100;
        ws[3, 1].Value = 2;
        ws[3, 2].Value = 250;

        var table = ws.Tables.Add("A1:B3", "SalesTable");
        table.Name.Should().Be("SalesTable");
        table.Columns.Count.Should().Be(2);
        table.Columns[0].Name.Should().Be("ID");
        table.Columns[1].Name.Should().Be("Amount");

        table.Columns[1].TotalsRowFunction = TotalsRowFunction.Sum;
        table.Columns[1].TotalsRowFunction.Should().Be(TotalsRowFunction.Sum);
    }

    [Fact]
    public void AutoFilterAndSort_WorksAsExpected()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("SortSheet");

        ws[1, 1].Value = 30;
        ws[2, 1].Value = 10;
        ws[3, 1].Value = 20;

        ws.AutoFilter("A1:A3");
        ws.AutoFilterRange.Should().NotBeNull();

        ws.Sort(1, 1, 3, 1, 1, ascending: true);

        ws[1, 1].GetValue<int>().Should().Be(10);
        ws[2, 1].GetValue<int>().Should().Be(20);
        ws[3, 1].GetValue<int>().Should().Be(30);
    }
}
