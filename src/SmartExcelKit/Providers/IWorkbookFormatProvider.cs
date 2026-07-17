namespace SmartExcelKit.Providers;

/// <summary>
/// Defines the layout contract for reading and writing workbooks in specific file formats.
/// </summary>
public interface IWorkbookFormatProvider
{
    /// <summary>
    /// Reads workbook contents from a stream.
    /// </summary>
    void Read(Stream stream, ExcelWorkbook workbook);

    /// <summary>
    /// Writes workbook contents to a stream.
    /// </summary>
    void Write(Stream stream, ExcelWorkbook workbook);

    /// <summary>
    /// Asynchronously reads workbook contents from a stream.
    /// </summary>
    Task ReadAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken);

    /// <summary>
    /// Asynchronously writes workbook contents to a stream.
    /// </summary>
    Task WriteAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken);
}
