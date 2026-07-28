using FluentAssertions;
using SmartExcelKit.Core;
using Xunit;

namespace SmartExcelKit.Tests;

/// <summary>
/// Unit tests for modern API features including indexers, used range, lazy enumerations, typed converters, and bulk operations.
/// </summary>
public class ModernApiTests
{
    /// <summary>Tests workbook indexers and ActiveWorksheet property.</summary>
    [Fact]
    public void Workbook_IndexersAndActiveWorksheet_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var sheet1 = workbook.AddWorksheet("Sheet1");
        var sheet2 = workbook.AddWorksheet("DataSheet");

        // Indexer by name
        workbook["Sheet1"].Should().BeSameAs(sheet1);
        workbook["DataSheet"].Should().BeSameAs(sheet2);

        // Indexer by index
        workbook[0].Should().BeSameAs(sheet1);
        workbook[1].Should().BeSameAs(sheet2);

        // ActiveWorksheet
        workbook.ActiveWorksheet.Should().BeSameAs(sheet1);
        workbook.ActiveWorksheet = sheet2;
        workbook.ActiveWorksheet.Should().BeSameAs(sheet2);
    }

    /// <summary>Tests worksheet indexers for single cell and range access.</summary>
    [Fact]
    public void Worksheet_IndexersAndNavigation_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        ws.Cell("A1").Value = "Header1";
        ws.Cell(1, 2).Value = "Header2";

        // Indexer cell access
        ws[1, 1].Value.Should().Be("Header1");
        ws[1, 2].Value.Should().Be("Header2");

        // Indexer range access
        var range = ws["A1:B1"];
        range.Address.Address.Should().Be("A1:B1");
        range.Cells.Count().Should().Be(2);
    }

    /// <summary>Tests UsedRange, row/col counts, and used boundary calculation.</summary>
    [Fact]
    public void Worksheet_UsedRangeAndUsedCounts_ShouldComputeAccurately()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        ws.UsedRange.Address.Address.Should().Be("A1");
        ws.UsedRowCount.Should().Be(0);
        ws.UsedColumnCount.Should().Be(0);
        ws.FirstRowUsed().Should().BeNull();
        ws.LastRowUsed().Should().BeNull();
        ws.FirstColumnUsed().Should().BeNull();
        ws.LastColumnUsed().Should().BeNull();

        // Populate B2 to D5
        ws["B2"].Value = 10;
        ws["D5"].Value = "End";

        ws.MaxRow.Should().Be(5);
        ws.MaxColumn.Should().Be(4);
        ws.RowCount.Should().Be(5);
        ws.ColumnCount.Should().Be(4);

        ws.UsedRange.Address.Address.Should().Be("B2:D5");
        ws.UsedRowCount.Should().Be(4); // rows 2, 3, 4, 5
        ws.UsedColumnCount.Should().Be(3); // cols B, C, D

        ws.FirstRowUsed()!.Index.Should().Be(2);
        ws.LastRowUsed()!.Index.Should().Be(5);
        ws.FirstColumnUsed()!.Index.Should().Be(2);
        ws.LastColumnUsed()!.Index.Should().Be(4);

        ws.CellsUsed().Count().Should().Be(2);
        ws.RowsUsed().Select(r => r.Index).Should().Equal(2, 5);
        ws.ColumnsUsed().Select(c => c.Index).Should().Equal(2, 4);
    }

    /// <summary>Tests ExcelRow properties, height, visibility, and cell access.</summary>
    [Fact]
    public void Row_PropertiesAndCells_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        var row = ws.Row(3);
        row.Index.Should().Be(3);
        row.RowNumber.Should().Be(3);
        row.IsEmpty.Should().BeTrue();

        row.Height = 25.5;
        ws.GetRowHeight(3).Should().Be(25.5);

        row.Hidden = true;
        ws.IsRowHidden(3).Should().BeTrue();

        row[1].Value = "Data";
        row.IsEmpty.Should().BeFalse();
        row.CellsUsed().Count().Should().Be(1);
    }

    /// <summary>Tests ExcelColumn properties, width, letter conversion, and cell access.</summary>
    [Fact]
    public void Column_PropertiesAndCells_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        var col = ws.Column(2);
        col.Index.Should().Be(2);
        col.ColumnNumber.Should().Be(2);
        col.Letter.Should().Be("B");
        col.IsEmpty.Should().BeTrue();

        col.Width = 18.0;
        ws.GetColumnWidth(2).Should().Be(18.0);

        col.Hidden = true;
        ws.IsColumnHidden(2).Should().BeTrue();

        col[2].Value = "Column Data";
        col.IsEmpty.Should().BeFalse();
        col.CellsUsed().Count().Should().Be(1);
    }

    /// <summary>Tests typed GetValue, TryGetValue, and primitive type getters on ExcelCell.</summary>
    [Fact]
    public void Cell_TypedGettersAndConverters_ShouldBeRobust()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        var cellInt = ws.Cell("A1");
        cellInt.Value = 123;
        cellInt.HasValue.Should().BeTrue();
        cellInt.RowNumber.Should().Be(1);
        cellInt.ColumnNumber.Should().Be(1);
        cellInt.GetInt32().Should().Be(123);
        cellInt.GetInt64().Should().Be(123L);
        cellInt.GetDouble().Should().Be(123.0);
        cellInt.GetDecimal().Should().Be(123m);
        cellInt.GetValue<int>().Should().Be(123);
        cellInt.TryGetValue<int>(out int intVal).Should().BeTrue();
        intVal.Should().Be(123);

        var cellBool = ws.Cell("A2");
        cellBool.Value = "true";
        cellBool.GetBoolean().Should().BeTrue();
        cellBool.GetValue<bool>().Should().BeTrue();

        var cellDate = ws.Cell("A3");
        var dt = new DateTime(2026, 7, 27, 10, 0, 0);
        cellDate.Value = dt;
        cellDate.GetDateTime().Should().Be(dt);
        cellDate.GetValue<DateTime>().Should().Be(dt);

        var emptyCell = ws.Cell("Z99");
        emptyCell.HasValue.Should().BeFalse();
        emptyCell.GetString().Should().BeEmpty();
        emptyCell.TryGetValue<int>(out _).Should().BeFalse();
    }

    /// <summary>Tests bulk matrix SetValues, GetValues, CopyTo, and Clear operations on ExcelRange.</summary>
    [Fact]
    public void Range_BulkOperationsAndCopyTo_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        var range = ws.Range("A1:C3");
        range.Rows.Count().Should().Be(3);
        range.Columns.Count().Should().Be(3);
        range.Cells.Count().Should().Be(9);

        // Bulk 2D values
        var matrix = new object?[,]
        {
            { 1, 2, 3 },
            { 4, 5, 6 },
            { 7, 8, 9 }
        };
        range.SetValues(matrix);

        range.GetValues()[1, 1].Should().Be(5);
        ws["B2"].Value.Should().Be(5);

        // CopyTo another range
        var targetRange = ws.Range("E1:G3");
        range.CopyTo(targetRange);

        ws["F2"].Value.Should().Be(5);

        // Clear contents
        range.ClearContents();
        ws["A1"].Value.Should().BeNull();
        targetRange["F2"].Value.Should().Be(5); // target unchanged
    }

    /// <summary>Tests inserting and deleting rows and columns with cell coordinate shifting.</summary>
    [Fact]
    public void Worksheet_InsertAndDeleteRowsColumns_ShouldShiftCellsCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        ws["A1"].Value = "R1C1";
        ws["A2"].Value = "R2C1";
        ws["B1"].Value = "R1C2";

        // Insert row at 2
        ws.InsertRows(2, 1);
        ws["A1"].Value.Should().Be("R1C1");
        ws["A2"].Value.Should().BeNull(); // inserted blank row
        ws["A3"].Value.Should().Be("R2C1"); // shifted down

        // Delete row at 2
        ws.DeleteRows(2, 1);
        ws["A1"].Value.Should().Be("R1C1");
        ws["A2"].Value.Should().Be("R2C1"); // shifted back up

        // Insert col at 2
        ws.InsertColumns(2, 1);
        ws["A1"].Value.Should().Be("R1C1");
        ws["B1"].Value.Should().BeNull();
        ws["C1"].Value.Should().Be("R1C2");

        // Delete col at 2
        ws.DeleteColumns(2, 1);
        ws["A1"].Value.Should().Be("R1C1");
        ws["B1"].Value.Should().Be("R1C2");
    }

    /// <summary>Tests inserting and deleting rows with merged range shifting.</summary>
    [Fact]
    public void Worksheet_InsertAndDeleteRows_ShouldShiftMergedRangesCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        var range = ExcelRangeAddress.Parse("A5:B6");
        ws.MergeCells(range);
        ws.MergedRanges.Count.Should().Be(1);

        // Insert 2 rows at row 2
        ws.InsertRows(2, 2);
        ws.MergedRanges[0].Address.Should().Be("A7:B8");

        // Delete 2 rows at row 2
        ws.DeleteRows(2, 2);
        ws.MergedRanges[0].Address.Should().Be("A5:B6");
    }

    /// <summary>Tests zero-allocation struct enumerators on Range, Row, and Column.</summary>
    [Fact]
    public void StructEnumerators_ShouldIterateWithoutAllocations()
    {
        using var workbook = new ExcelWorkbook();
        var ws = workbook.AddWorksheet("Test");

        ws["A1"].Value = 1;
        ws["B2"].Value = 2;

        int cellCount = 0;
        foreach (ExcelCell cell in ws.Range("A1:B2"))
        {
            cellCount++;
        }
        cellCount.Should().Be(4);

        int rowCellCount = 0;
        foreach (ExcelCell cell in ws.Row(1))
        {
            rowCellCount++;
        }
        rowCellCount.Should().Be(2); // up to max col 2

        int colCellCount = 0;
        foreach (ExcelCell cell in ws.Column(1))
        {
            colCellCount++;
        }
        colCellCount.Should().Be(2); // up to max row 2
    }
}
