using SmartExcelKit.Csv;
using SmartExcelKit.Exceptions;
using System.Text;

namespace SmartExcelKit.Providers;

/// <summary>
/// Format provider for reading and writing CSV/TSV files using CsvEngine.
/// </summary>
public sealed class CsvFormatProvider : IWorkbookFormatProvider
{
    private readonly char _delimiter;
    private readonly Encoding? _encoding;

    /// <summary>
    /// Initializes a new instance of the <see cref="CsvFormatProvider"/> class.
    /// </summary>
    /// <param name="delimiter">The delimiter char (e.g. ',' for CSV, '\t' for TSV).</param>
    /// <param name="encoding">The encoding to use (auto-detected if null during read).</param>
    public CsvFormatProvider(char delimiter = ',', Encoding? encoding = null)
    {
        _delimiter = delimiter;
        _encoding = encoding;
    }

    /// <inheritdoc />
    public void Read(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        // Clear existing sheets
        while (workbook.Worksheets.Count > 0)
        {
            workbook.RemoveWorksheet(workbook.Worksheets[0].Name);
        }

        var sheet = workbook.AddWorksheet("Sheet1");
        int row = 1;

        foreach (var fields in CsvEngine.ReadStreaming(stream, _delimiter, _encoding))
        {
            for (int col = 0; col < fields.Count; col++)
            {
                sheet.Cell(row, col + 1).Value = ParseValue(fields[col]);
            }
            row++;
        }
    }

    /// <inheritdoc />
    public void Write(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        if (workbook.Worksheets.Count == 0)
            throw new ExportException("Workbook must contain at least one worksheet to export to CSV.", "NO_WORKSHEET_TO_EXPORT");

        var sheet = workbook.Worksheets[0];
        int maxRow = sheet.MaxRow;
        int maxCol = sheet.MaxColumn;

        var rowsList = new List<List<string>>();

        for (int r = 1; r <= maxRow; r++)
        {
            var rowFields = new List<string>();
            for (int c = 1; c <= maxCol; c++)
            {
                rowFields.Add(sheet.Cell(r, c).GetString());
            }
            rowsList.Add(rowFields);
        }

        CsvEngine.Write(stream, rowsList, _delimiter, _encoding);
    }

    /// <inheritdoc />
    public Task ReadAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        // CsvEngine.ReadStreaming is currently synchronous but wraps reading from streams. We run it in Task context for non-blocking.
        return Task.Run(() => Read(stream, workbook), cancellationToken);
    }

    /// <inheritdoc />
    public Task WriteAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(stream, workbook), cancellationToken);
    }

    private static object ParseValue(string raw)
    {
        if (double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double d)) return d;
        if (bool.TryParse(raw, out bool b)) return b;
        return raw;
    }
}
