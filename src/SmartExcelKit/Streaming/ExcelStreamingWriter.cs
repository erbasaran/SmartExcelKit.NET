using SmartExcelKit.Core;
using System.IO.Compression;
using System.Text;

namespace SmartExcelKit.Streaming;

/// <summary>
/// A high-performance, forward-only streaming XLSX writer.
/// Designed for low memory usage when writing very large datasets.
/// </summary>
public sealed class ExcelStreamingWriter : IDisposable
{
    private readonly Stream _outputStream;
    private readonly ZipArchive _archive;
    private readonly List<(string Name, int SheetId, string Path)> _sheets = [];

    private ZipArchiveEntry? _currentEntry;
    private StreamWriter? _currentWriter;
    private int _currentRow;
    private bool _isDisposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelStreamingWriter"/> class wrapping a stream.
    /// </summary>
    /// <param name="outputStream">The target output stream.</param>
    public ExcelStreamingWriter(Stream outputStream)
    {
        _outputStream = outputStream ?? throw new ArgumentNullException(nameof(outputStream));
        _archive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true);
    }

    /// <summary>
    /// Begins a new worksheet in the ZIP package. Must be called before writing rows.
    /// </summary>
    /// <param name="sheetName">The name of the worksheet.</param>
    public void BeginSheet(string sheetName)
    {
        if (_isDisposed) throw new ObjectDisposedException(nameof(ExcelStreamingWriter));
        if (string.IsNullOrWhiteSpace(sheetName)) throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));

        EndSheet();

        int sheetId = _sheets.Count + 1;
        string wsPath = $"worksheets/sheet{sheetId}.xml";
        _sheets.Add((sheetName, sheetId, wsPath));

        _currentEntry = _archive.CreateEntry($"xl/{wsPath}");
        var entryStream = _currentEntry.Open();
        _currentWriter = new StreamWriter(entryStream, Encoding.UTF8);
        _currentRow = 1;

        // Write worksheet header
        _currentWriter.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
        _currentWriter.Write("<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
        _currentWriter.Write("<sheetData>");
    }

    /// <summary>
    /// Writes a single row of cell values to the currently active worksheet.
    /// </summary>
    /// <param name="values">The collection of cell values to write.</param>
    public void WriteRow(IEnumerable<object?> values)
    {
        if (_isDisposed) throw new ObjectDisposedException(nameof(ExcelStreamingWriter));
        if (_currentWriter == null) throw new InvalidOperationException("No active sheet. Call BeginSheet() first.");

        _currentWriter.Write($"<row r=\"{_currentRow}\">");

        int colIndex = 1;
        foreach (var val in values)
        {
            if (val != null)
            {
                string colName = CellAddress.GetColumnName(colIndex);
                string cellRef = $"{colName}{_currentRow}";

                if (val is string s)
                {
                    // Escape XML characters
                    string escaped = System.Security.SecurityElement.Escape(s);
                    // Use inlineStr to avoid building a Shared String Table in memory
                    _currentWriter.Write($"<c r=\"{cellRef}\" t=\"inlineStr\"><is><t>{escaped}</t></is></c>");
                }
                else if (val is bool b)
                {
                    _currentWriter.Write($"<c r=\"{cellRef}\" t=\"b\"><v>{(b ? "1" : "0")}</v></c>");
                }
                else if (val is double d)
                {
                    _currentWriter.Write($"<c r=\"{cellRef}\"><v>{d.ToString(System.Globalization.CultureInfo.InvariantCulture)}</v></c>");
                }
                else if (val is float f)
                {
                    _currentWriter.Write($"<c r=\"{cellRef}\"><v>{f.ToString(System.Globalization.CultureInfo.InvariantCulture)}</v></c>");
                }
                else if (val is decimal dec)
                {
                    _currentWriter.Write($"<c r=\"{cellRef}\"><v>{dec.ToString(System.Globalization.CultureInfo.InvariantCulture)}</v></c>");
                }
                else if (val is int i)
                {
                    _currentWriter.Write($"<c r=\"{cellRef}\"><v>{i}</v></c>");
                }
                else if (val is long l)
                {
                    _currentWriter.Write($"<c r=\"{cellRef}\"><v>{l}</v></c>");
                }
                else if (val is DateTime dt)
                {
                    _currentWriter.Write($"<c r=\"{cellRef}\"><v>{dt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture)}</v></c>");
                }
                else
                {
                    string escapedVal = System.Security.SecurityElement.Escape(val.ToString() ?? string.Empty);
                    _currentWriter.Write($"<c r=\"{cellRef}\" t=\"inlineStr\"><is><t>{escapedVal}</t></is></c>");
                }
            }
            colIndex++;
        }

        _currentWriter.Write("</row>");
        _currentRow++;
    }

    /// <summary>
    /// Closes the currently active worksheet.
    /// </summary>
    public void EndSheet()
    {
        if (_currentWriter != null)
        {
            _currentWriter.Write("</sheetData>");
            _currentWriter.Write("</worksheet>");
            _currentWriter.Flush();
            _currentWriter.Dispose();
            _currentWriter = null;
            _currentEntry = null;
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (!_isDisposed)
        {
            EndSheet();
            WriteMetadataFiles();
            _archive.Dispose();
            _isDisposed = true;
        }
    }

    private void WriteMetadataFiles()
    {
        // 1. Write xl/styles.xml (Minimal layout to satisfy Excel parser)
        var stylesEntry = _archive.CreateEntry("xl/styles.xml");
        using (var stream = stylesEntry.Open())
        using (var writer = new StreamWriter(stream, Encoding.UTF8))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">");
            writer.Write("<fonts count=\"1\"><font><sz val=\"11\"/><name val=\"Calibri\"/><family val=\"2\"/></font></fonts>");
            writer.Write("<fills count=\"2\"><fill><patternFill patternType=\"none\"/></fill><fill><patternFill patternType=\"gray125\"/></fill></fills>");
            writer.Write("<borders count=\"1\"><border><left/><right/><top/><bottom/></border></borders>");
            writer.Write("<cellStyleXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\"/></cellStyleXfs>");
            writer.Write("<cellXfs count=\"1\"><xf numFmtId=\"0\" fontId=\"0\" fillId=\"0\" borderId=\"0\" xfId=\"0\"/></cellXfs>");
            writer.Write("</styleSheet>");
        }

        // 2. Write xl/_rels/workbook.xml.rels
        var wbRelsEntry = _archive.CreateEntry("xl/_rels/workbook.xml.rels");
        using (var stream = wbRelsEntry.Open())
        using (var writer = new StreamWriter(stream, Encoding.UTF8))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            foreach (var sheet in _sheets)
            {
                writer.Write($"<Relationship Id=\"rId{sheet.SheetId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"{sheet.Path}\" />");
            }
            writer.Write($"<Relationship Id=\"rIdStyles\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" Target=\"styles.xml\" />");
            writer.Write("</Relationships>");
        }

        // 3. Write xl/workbook.xml
        var wbEntry = _archive.CreateEntry("xl/workbook.xml");
        using (var stream = wbEntry.Open())
        using (var writer = new StreamWriter(stream, Encoding.UTF8))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">");
            writer.Write("<sheets>");
            foreach (var sheet in _sheets)
            {
                writer.Write($"<sheet name=\"{sheet.Name}\" sheetId=\"{sheet.SheetId}\" r:id=\"rId{sheet.SheetId}\" />");
            }
            writer.Write("</sheets>");
            writer.Write("</workbook>");
        }

        // 4. Write _rels/.rels
        var rootRelsEntry = _archive.CreateEntry("_rels/.rels");
        using (var stream = rootRelsEntry.Open())
        using (var writer = new StreamWriter(stream, Encoding.UTF8))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">");
            writer.Write("<Relationship Id=\"rIdWb\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\" />");
            writer.Write("</Relationships>");
        }

        // 5. Write [Content_Types].xml
        var ctEntry = _archive.CreateEntry("[Content_Types].xml");
        using (var stream = ctEntry.Open())
        using (var writer = new StreamWriter(stream, Encoding.UTF8))
        {
            writer.Write("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
            writer.Write("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">");
            writer.Write("<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\" />");
            writer.Write("<Default Extension=\"xml\" ContentType=\"application/xml\" />");
            writer.Write("<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\" />");
            writer.Write("<Override PartName=\"/xl/styles.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\" />");
            foreach (var sheet in _sheets)
            {
                writer.Write($"<Override PartName=\"/xl/{sheet.Path}\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\" />");
            }
            writer.Write("</Types>");
        }
    }
}
