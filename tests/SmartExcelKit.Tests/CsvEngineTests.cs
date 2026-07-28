using System.Text;
using FluentAssertions;
using SmartExcelKit.Csv;
using Xunit;

namespace SmartExcelKit.Tests;

/// <summary>
/// Unit tests for CsvEngine parsing, quoted fields, escaped quotes, and delimiter detection.
/// </summary>
public class CsvEngineTests
{
    /// <summary>Tests parsing simple CSV datasets.</summary>
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

    /// <summary>Tests parsing quoted fields containing newlines.</summary>
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

    /// <summary>Tests parsing escaped quotes inside quoted fields.</summary>
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

    /// <summary>Tests auto-detecting CSV delimiters (semicolon, comma, tab).</summary>
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
