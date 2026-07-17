using FluentAssertions;
using SmartExcelKit.Csv;
using System.Text;
using Xunit;

namespace SmartExcelKit.Tests;

public class CsvEngineTests
{
    [Fact]
    public void CsvEngine_ShouldParseSimpleCsv()
    {
        // Arrange
        string csv = "Name,Age,City\nAlice,30,New York\nBob,25,London";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        // Act
        var rows = CsvEngine.Read(stream, ',');

        // Assert
        rows.Count.Should().Be(3);
        rows[0][0].Should().Be("Name");
        rows[1][1].Should().Be("30");
        rows[2][2].Should().Be("London");
    }

    [Fact]
    public void CsvEngine_ShouldParseQuotedFieldsWithNewlines()
    {
        // Arrange
        string csv = "ID,Message\n1,\"Hello\nWorld\"\n2,\"Normal message\"";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        // Act
        var rows = CsvEngine.Read(stream, ',');

        // Assert
        rows.Count.Should().Be(3);
        rows[1][1].Should().Be("Hello\nWorld");
        rows[2][1].Should().Be("Normal message");
    }

    [Fact]
    public void CsvEngine_ShouldParseEscapedQuotes()
    {
        // Arrange
        string csv = "ID,Name\n1,\"Alice \"\"The Queen\"\"\"";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        // Act
        var rows = CsvEngine.Read(stream, ',');

        // Assert
        rows.Count.Should().Be(2);
        rows[1][1].Should().Be("Alice \"The Queen\"");
    }

    [Fact]
    public void CsvEngine_ShouldDetectDelimiter()
    {
        // Arrange
        string csv = "Name;Age;City\nAlice;30;Paris";
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(csv));

        // Act
        char delim = CsvEngine.DetectDelimiter(stream, Encoding.UTF8);

        // Assert
        delim.Should().Be(';');
    }
}
