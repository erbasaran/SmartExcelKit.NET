using System.Data;
using FluentAssertions;
using Xunit;

namespace SmartExcelKit.Tests;

/// <summary>
/// Unit tests for developer convenience helper APIs across Cell, Row, Range, and Worksheet.
/// </summary>
public class HelperApiTests
{
    /// <summary>Sample DTO for helper mapping tests.</summary>
    public class PersonDto
    {
        /// <summary>Gets or sets person name.</summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>Gets or sets person age.</summary>
        public int Age { get; set; }

        /// <summary>Gets or sets person salary.</summary>
        public double Salary { get; set; }
    }

    /// <summary>Tests Cell GetValueOrDefault and fluid getters.</summary>
    [Fact]
    public void CellHelpers_GetValueOrDefaultAndFluidGetters_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Test");

        var cell = sheet.Cell("A1");
        cell.Value = 42;

        cell.GetValueOrDefault(100).Should().Be(42);
        cell.AsString().Should().Be("42");
        cell.AsInt32().Should().Be(42);
        cell.AsInt64().Should().Be(42L);
        cell.AsDouble().Should().Be(42.0);
        cell.AsDecimal().Should().Be(42m);

        var emptyCell = sheet.Cell("B1");
        emptyCell.GetValueOrDefault(999).Should().Be(999);
        emptyCell.GetValueOrDefault("Default").Should().Be("Default");
    }

    /// <summary>Tests Row helpers (Values, ToDictionary, ToCsv, ToJson, ToObject, IsBlank, FirstCellUsed, LastCellUsed).</summary>
    [Fact]
    public void RowHelpers_ValuesToDictionaryToCsvToJsonToObject_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Test");

        // Header row
        sheet["A1"].Value = "Name";
        sheet["B1"].Value = "Age";
        sheet["C1"].Value = "Salary";

        // Data row
        var row2 = sheet.Row(2);
        row2[1].Value = "Alice";
        row2[2].Value = 30;
        row2[3].Value = 95000.50;

        row2.Values().Should().Equal("Alice", 30, 95000.50);
        row2.ToArray().Should().Equal("Alice", 30, 95000.50);
        row2.ToList().Should().Equal("Alice", 30, 95000.50);

        // ToCsv
        string csvLine = row2.ToCsv(delimiter: ';');
        csvLine.Should().Be("Alice;30;95000.5");

        // ToDictionary
        var dict = row2.ToDictionary(headerRow: 1);
        dict["Name"].Should().Be("Alice");
        dict["Age"].Should().Be(30);
        dict["Salary"].Should().Be(95000.50);

        // ToJson
        string json = row2.ToJson(headerRow: 1);
        json.Should().Contain("\"Name\":\"Alice\"");

        // ToObject<T>
        var person = row2.ToObject<PersonDto>(headerRow: 1);
        person.Name.Should().Be("Alice");
        person.Age.Should().Be(30);
        person.Salary.Should().Be(95000.50);

        // FirstCellUsed and LastCellUsed
        row2.FirstCellUsed()!.Address.Address.Should().Be("A2");
        row2.LastCellUsed()!.Address.Address.Should().Be("C2");
        row2.IsBlank().Should().BeFalse();

        // Blank row test
        var blankRow = sheet.Row(10);
        blankRow.IsBlank().Should().BeTrue();
        blankRow.FirstCellUsed().Should().BeNull();
        blankRow.LastCellUsed().Should().BeNull();
    }

    /// <summary>Tests Range helpers (ToCsv, ToJson, ToObjects, ToDataTable, ToDictionaryList).</summary>
    [Fact]
    public void RangeHelpers_ToCsvToJsonToObjectsToDataTable_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Test");

        sheet["A1"].Value = "Name";
        sheet["B1"].Value = "Age";
        sheet["A2"].Value = "Bob";
        sheet["B2"].Value = 25;
        sheet["A3"].Value = "Charlie";
        sheet["B3"].Value = 35;

        var range = sheet.Range("A1:B3");

        // ToArray / ToList 1D flattened
        range.ToArray().Length.Should().Be(6);
        range.ToList().Count.Should().Be(6);

        // ToCsv
        string csv = range.ToCsv();
        csv.Should().Contain("Name,Age");
        csv.Should().Contain("Bob,25");
        csv.Should().Contain("Charlie,35");

        // ToDictionaryList
        var dictList = range.ToDictionaryList(hasHeader: true);
        dictList.Count.Should().Be(2);
        dictList[0]["Name"].Should().Be("Bob");
        dictList[1]["Name"].Should().Be("Charlie");

        // ToJson
        string json = range.ToJson(hasHeader: true);
        json.Should().Contain("\"Name\":\"Bob\"");

        // ToDataTable
        DataTable dt = range.ToDataTable(hasHeader: true);
        dt.Rows.Count.Should().Be(2);
        dt.Columns.Count.Should().Be(2);

        // ToObjects
        var people = range.ToObjects<PersonDto>(hasHeader: true).ToList();
        people.Count.Should().Be(2);
        people[0].Name.Should().Be("Bob");
        people[1].Name.Should().Be("Charlie");
    }

    /// <summary>Tests Worksheet helpers (ToCsv, ToJson, ToDataTable, ToObjects, ToDictionaryList).</summary>
    [Fact]
    public void WorksheetHelpers_ToCsvToJsonToDataTableToObjects_ShouldWorkCorrectly()
    {
        using var workbook = new ExcelWorkbook();
        var sheet = workbook.AddWorksheet("Employees");

        sheet["A1"].Value = "Name";
        sheet["B1"].Value = "Age";
        sheet["A2"].Value = "Dave";
        sheet["B2"].Value = 40;

        // ToCsv
        string csv = sheet.ToCsv(delimiter: ',');
        csv.Should().Contain("Name,Age");
        csv.Should().Contain("Dave,40");

        // ToJson
        string json = sheet.ToJson(hasHeader: true);
        json.Should().Contain("\"Name\":\"Dave\"");

        // ToDataTable
        DataTable dt = sheet.ToDataTable(hasHeader: true);
        dt.Rows.Count.Should().Be(1);
        dt.Rows[0]["Name"].Should().Be("Dave");

        // ToObjects
        var list = sheet.ToObjects<PersonDto>().ToList();
        list.Count.Should().Be(1);
        list[0].Name.Should().Be("Dave");

        // ToDictionaryList
        var dicts = sheet.ToDictionaryList();
        dicts.Count.Should().Be(1);
        dicts[0]["Name"].Should().Be("Dave");
    }
}
