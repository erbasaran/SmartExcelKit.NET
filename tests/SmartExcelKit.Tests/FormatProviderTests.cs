using FluentAssertions;
using Xunit;

namespace SmartExcelKit.Tests;

public class FormatProviderTests
{
    [Fact]
    public void XlsxProvider_ShouldWriteAndReadBackSymmetrically()
    {
        // Arrange
        using var originalWb = new ExcelWorkbook();
        var sheet = originalWb.AddWorksheet("TestSheet");
        sheet.Cell("A1").Value = "Product Name";
        sheet.Cell("B1").Value = 99.99;
        sheet.Cell("C1").Formula = "B1*2";

        using var ms = new MemoryStream();

        // Act - Save to memory stream
        originalWb.Save(ms, ExcelFileFormat.Xlsx);
        ms.Position = 0;

        // Act - Load back
        using var loadedWb = ExcelWorkbook.Open(ms);

        // Assert
        loadedWb.Worksheets.Count.Should().Be(1);
        var loadedSheet = loadedWb.Worksheets[0];
        loadedSheet.Name.Should().Be("TestSheet");
        loadedSheet.Cell("A1").Value.Should().Be("Product Name");
        loadedSheet.Cell("B1").Value.Should().Be(99.99);
        loadedSheet.Cell("C1").Formula.Should().Be("B1*2");
    }

    [Fact]
    public void JsonProvider_ShouldWriteAndReadBackSymmetrically()
    {
        // Arrange
        using var originalWb = new ExcelWorkbook();
        var sheet = originalWb.AddWorksheet("Data");
        sheet.Cell("A1").Value = "Value1";
        sheet.Cell("B1").Value = 123.45;
        sheet.Cell("C1").Value = true;

        using var ms = new MemoryStream();

        // Act - Save
        originalWb.Save(ms, ExcelFileFormat.Json);
        ms.Position = 0;

        // Act - Load
        using var loadedWb = ExcelWorkbook.Open(ms);

        // Assert
        var loadedSheet = loadedWb.Worksheets[0];
        loadedSheet.Cell("A1").Value.Should().Be("Value1");
        loadedSheet.Cell("B1").Value.Should().Be(123.45);
        loadedSheet.Cell("C1").Value.Should().Be(true);
    }

    [Fact]
    public void Xml2003Provider_ShouldWriteAndReadBackSymmetrically()
    {
        // Arrange
        using var originalWb = new ExcelWorkbook();
        var sheet = originalWb.AddWorksheet("XmlSheet");
        sheet.Cell("A1").Value = "Hello XML";
        sheet.Cell("B1").Value = 456.78;

        using var ms = new MemoryStream();

        // Act
        originalWb.Save(ms, ExcelFileFormat.Xml2003);
        ms.Position = 0;

        using var loadedWb = ExcelWorkbook.Open(ms);

        // Assert
        var loadedSheet = loadedWb.Worksheets[0];
        loadedSheet.Cell("A1").Value.Should().Be("Hello XML");
        loadedSheet.Cell("B1").Value.Should().Be(456.78);
    }

    [Fact]
    public void HtmlTableProvider_ShouldWriteAndReadBackSymmetrically()
    {
        // Arrange
        using var originalWb = new ExcelWorkbook();
        var sheet = originalWb.AddWorksheet("HtmlSheet");
        sheet.Cell("A1").Value = "HTML Content";
        sheet.Cell("B1").Value = 789.0;

        using var ms = new MemoryStream();

        // Act
        originalWb.Save(ms, ExcelFileFormat.HtmlTable);
        ms.Position = 0;

        using var loadedWb = ExcelWorkbook.Open(ms);

        // Assert
        var loadedSheet = loadedWb.Worksheets[0];
        loadedSheet.Cell("A1").Value.Should().Be("HTML Content");
        loadedSheet.Cell("B1").Value.Should().Be(789.0);
    }
}
