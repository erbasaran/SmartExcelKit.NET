using FluentAssertions;
using SmartExcelKit.Streaming;
using Xunit;

namespace SmartExcelKit.Tests;

public class StreamingTests
{
    [Fact]
    public void StreamingWriterAndReader_ShouldWriteAndReadSymmetrically()
    {
        // Arrange
        using var ms = new MemoryStream();

        // Act: Write using streaming writer
        using (var writer = new ExcelStreamingWriter(ms))
        {
            writer.BeginSheet("StreamingSheet");

            writer.WriteRow(new object?[] { "ColA", "ColB", "ColC" });
            writer.WriteRow(new object?[] { 456.78, true, "Value C" });
            writer.WriteRow(new object?[] { null, "B3 Value" }); // Row 3 with sparse column A
        }

        ms.Position = 0;

        // Act: Read using streaming reader
        using var reader = new ExcelStreamingReader(ms);
        var sheets = reader.GetSheets().ToList();

        // Assert
        sheets.Should().ContainSingle().Which.Should().Be("StreamingSheet");

        var rows = reader.ReadRows("StreamingSheet").ToList();
        rows.Count.Should().Be(3);

        // Row 1 assertions
        rows[0].Length.Should().Be(3);
        rows[0][0].Should().Be("ColA");
        rows[0][1].Should().Be("ColB");
        rows[0][2].Should().Be("ColC");

        // Row 2 assertions
        rows[1].Length.Should().Be(3);
        rows[1][0].Should().Be(456.78);
        rows[1][1].Should().Be(true);
        rows[1][2].Should().Be("Value C");

        // Row 3 assertions (Checks cell padding alignment)
        rows[2].Length.Should().Be(2);
        rows[2][0].Should().BeNull(); // padded column A
        rows[2][1].Should().Be("B3 Value");
    }
}
