using System.Data;
using FluentAssertions;
using Xunit;

namespace SmartExcelKit.Tests;

public class WorkbookExportTests
{
    public class SampleProduct
    {
        public string Name { get; set; } = string.Empty;
        public double Price { get; set; }
    }

    [Fact]
    public void Workbook_ExportMethods_ReturnAggregatedDataAcrossSheets()
    {
        using var workbook = new ExcelWorkbook();
        var sheet1 = workbook.AddWorksheet("Electronics");
        sheet1[1, 1].Value = "Name";
        sheet1[1, 2].Value = "Price";
        sheet1[2, 1].Value = "Laptop";
        sheet1[2, 2].Value = 999.99;

        var sheet2 = workbook.AddWorksheet("Accessories");
        sheet2[1, 1].Value = "Name";
        sheet2[1, 2].Value = "Price";
        sheet2[2, 1].Value = "Mouse";
        sheet2[2, 2].Value = 29.99;

        // 1. ToCsv across workbook
        string csv = workbook.ToCsv();
        csv.Should().Contain("Laptop");
        csv.Should().Contain("Mouse");

        // 2. ToJson across workbook
        string json = workbook.ToJson();
        json.Should().Contain("Electronics");
        json.Should().Contain("Accessories");

        // 3. ToDataSet across workbook
        DataSet ds = workbook.ToDataSet();
        ds.Tables.Count.Should().Be(2);
        ds.Tables["Electronics"].Should().NotBeNull();
        ds.Tables["Accessories"].Should().NotBeNull();

        // 4. ToObjects across workbook
        var products = workbook.ToObjects<SampleProduct>().ToList();
        products.Count.Should().Be(2);
        products.Select(p => p.Name).Should().Contain(new[] { "Laptop", "Mouse" });

        // 5. ToDictionaryList across workbook
        var dictList = workbook.ToDictionaryList();
        dictList.Count.Should().Be(2);
        dictList[0]["__Worksheet"].Should().Be("Electronics");
        dictList[1]["__Worksheet"].Should().Be("Accessories");
    }
}
