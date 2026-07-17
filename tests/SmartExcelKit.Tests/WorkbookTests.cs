using FluentAssertions;
using SmartExcelKit.Core;
using Xunit;

namespace SmartExcelKit.Tests;

public class WorkbookTests
{
    [Fact]
    public void Workbook_ShouldCreateEmptySheetWithCorrectName()
    {
        // Arrange & Act
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("TestSheet");

        // Assert
        sheet.Should().NotBeNull();
        sheet.Name.Should().Be("TestSheet");
        workbook.Worksheets.Count.Should().Be(1);
    }

    [Fact]
    public void Workbook_ShouldThrowOnDuplicateWorksheetName()
    {
        // Arrange
        using var workbook = new ExcelWorkbook();
        workbook.AddWorksheet("TestSheet");

        // Act
        var act = () => workbook.AddWorksheet("testsheet");

        // Assert
        act.Should().Throw<Exception>();
    }

    [Fact]
    public void CellAddress_ParseAndFormat_ShouldBeSymmetric()
    {
        // Arrange & Act
        var cell = CellAddress.Parse("AB12");

        // Assert
        cell.Column.Should().Be(28);
        cell.Row.Should().Be(12);
        cell.Address.Should().Be("AB12");
    }

    [Fact]
    public void RangeAddress_Contains_ShouldWorkCorrectly()
    {
        // Arrange
        var range = ExcelRangeAddress.Parse("B2:D4");

        // Act & Assert
        range.Contains(new CellAddress(2, 2)).Should().BeTrue(); // Top-Left
        range.Contains(new CellAddress(3, 3)).Should().BeTrue(); // Inside
        range.Contains(new CellAddress(4, 4)).Should().BeTrue(); // Bottom-Right
        range.Contains(new CellAddress(1, 1)).Should().BeFalse(); // Outside
        range.Contains(new CellAddress(5, 4)).Should().BeFalse(); // Outside
    }

    [Fact]
    public void Range_SetValue_ShouldApplyToAllCells()
    {
        // Arrange
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        var range = sheet.Range("A1:B2");

        // Act
        range.Value = 100.0;

        // Assert
        sheet.Cell("A1").Value.Should().Be(100.0);
        sheet.Cell("A2").Value.Should().Be(100.0);
        sheet.Cell("B1").Value.Should().Be(100.0);
        sheet.Cell("B2").Value.Should().Be(100.0);
    }

    [Fact]
    public void Sheet_ProtectAndUnprotect_ShouldWorkSymmetrically()
    {
        // Arrange
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");

        // Act & Assert
        sheet.IsProtected.Should().BeFalse();
        sheet.ProtectionPasswordHash.Should().BeNull();

        sheet.Protect("secret");
        sheet.IsProtected.Should().BeTrue();
        sheet.ProtectionPasswordHash.Should().NotBeNullOrEmpty();
        sheet.ProtectionPasswordHash.Should().Be("DAA7"); // standard excel XOR hash for 'secret'

        sheet.Unprotect();
        sheet.IsProtected.Should().BeFalse();
        sheet.ProtectionPasswordHash.Should().BeNull();
    }

    [Fact]
    public void Sheet_MergeAndUnmerge_ShouldTrackCorrectly()
    {
        // Arrange
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Sheet1");
        var range = ExcelRangeAddress.Parse("A1:C3");

        // Act
        sheet.MergeCells(range);

        // Assert
        sheet.MergedRanges.Count.Should().Be(1);
        sheet.MergedRanges[0].Address.Should().Be("A1:C3");

        // Act
        sheet.UnmergeCells(range);

        // Assert
        sheet.MergedRanges.Count.Should().Be(0);
    }
}

