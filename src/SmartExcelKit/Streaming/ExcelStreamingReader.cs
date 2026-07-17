using SmartExcelKit.Exceptions;
using System.IO.Compression;
using System.Xml;

namespace SmartExcelKit.Streaming;

/// <summary>
/// A high-performance, forward-only streaming XLSX reader.
/// Uses an XmlReader to yield rows one-by-one from zip streams.
/// </summary>
public sealed class ExcelStreamingReader : IDisposable
{
    private readonly Stream _inputStream;
    private readonly ZipArchive _archive;
    private readonly List<string> _sharedStrings = [];
    private readonly Dictionary<string, string> _sheetMap = []; // Name -> Entry Path
    private bool _isDisposed;

    /// <summary>
    /// Initializes a new instance of the <see cref="ExcelStreamingReader"/> class.
    /// </summary>
    /// <param name="inputStream">The input stream of the XLSX file.</param>
    public ExcelStreamingReader(Stream inputStream)
    {
        _inputStream = inputStream ?? throw new ArgumentNullException(nameof(inputStream));
        _archive = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: true);

        LoadSharedStrings();
        LoadWorkbookStructure();
    }

    /// <summary>
    /// Gets the names of all worksheets in the workbook.
    /// </summary>
    /// <returns>A collection of worksheet names.</returns>
    public IEnumerable<string> GetSheets() => _sheetMap.Keys;

    /// <summary>
    /// Reads and yields rows of cell values from the specified worksheet.
    /// Missing or empty cells are aligned and filled with null.
    /// </summary>
    /// <param name="sheetName">The name of the worksheet to read.</param>
    /// <returns>An enumerable of object arrays representing each row's values.</returns>
    public IEnumerable<object?[]> ReadRows(string sheetName)
    {
        if (_isDisposed) throw new ObjectDisposedException(nameof(ExcelStreamingReader));
        if (!_sheetMap.TryGetValue(sheetName, out string? entryPath))
        {
            throw new WorksheetException($"Worksheet '{sheetName}' not found in workbook.", "SHEET_NOT_FOUND");
        }

        var entry = _archive.GetEntry(entryPath);
        if (entry == null) yield break;

        using var stream = entry.Open();

        // Disable XML DTD/Schema resolution for security and speed
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true
        };

        using var reader = XmlReader.Create(stream, settings);

        var rowValues = new List<object?>();
        int lastRowIndex = 0;

        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "row")
            {
                rowValues.Clear();

                // Parse optional row index attribute
                string? rowAttr = reader.GetAttribute("r");
                if (int.TryParse(rowAttr, out int rIdx))
                {
                    lastRowIndex = rIdx;
                }
                else
                {
                    lastRowIndex++;
                }

                // Read child elements of the row
                if (!reader.IsEmptyElement)
                {
                    using var rowReader = reader.ReadSubtree();
                    rowReader.Read(); // Advance to <row>

                    while (rowReader.Read())
                    {
                        if (rowReader.NodeType == XmlNodeType.Element && rowReader.LocalName == "c")
                        {
                            string? cellRef = rowReader.GetAttribute("r");
                            string? type = rowReader.GetAttribute("t");

                            int colIndex = 1;
                            if (!string.IsNullOrEmpty(cellRef))
                            {
                                colIndex = GetColumnIndex(cellRef);
                            }

                            // Pad list if there are empty columns before this cell
                            while (rowValues.Count < colIndex)
                            {
                                rowValues.Add(null);
                            }

                            object? cellVal = null;
                            if (!rowReader.IsEmptyElement)
                            {
                                using var cellReader = rowReader.ReadSubtree();
                                cellReader.Read(); // Advance to <c>

                                while (cellReader.Read())
                                {
                                    if (cellReader.NodeType == XmlNodeType.Element && cellReader.LocalName == "v")
                                    {
                                        string rawVal = cellReader.ReadElementContentAsString();
                                        cellVal = ParseCellValue(rawVal, type);
                                    }
                                    else if (cellReader.NodeType == XmlNodeType.Element && cellReader.LocalName == "is")
                                    {
                                        // Inline string parsing <is><t>text</t></is>
                                        cellVal = ReadInlineString(cellReader);
                                    }
                                }
                            }

                            rowValues[colIndex - 1] = cellVal;
                        }
                    }
                }

                yield return rowValues.ToArray();
            }
        }
    }

    private void LoadSharedStrings()
    {
        var entry = _archive.GetEntry("xl/sharedStrings.xml");
        if (entry == null) return;

        using var stream = entry.Open();
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            IgnoreComments = true,
            IgnoreWhitespace = true
        };

        using var reader = XmlReader.Create(stream, settings);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "si")
            {
                using var siReader = reader.ReadSubtree();
                siReader.Read(); // Move to <si>
                string val = string.Empty;

                while (siReader.Read())
                {
                    if (siReader.NodeType == XmlNodeType.Element && siReader.LocalName == "t")
                    {
                        val = siReader.ReadElementContentAsString();
                    }
                }
                _sharedStrings.Add(val);
            }
        }
    }

    private void LoadWorkbookStructure()
    {
        var entry = _archive.GetEntry("xl/workbook.xml");
        if (entry == null) throw new ParsingException("Invalid XLSX package: workbook.xml not found.", "INVALID_XLSX");

        // Parse workbook relationships to map Sheet rId to Entry path
        var relsMap = new Dictionary<string, string>();
        var relsEntry = _archive.GetEntry("xl/_rels/workbook.xml.rels");
        if (relsEntry != null)
        {
            using var rStream = relsEntry.Open();
            using var rReader = XmlReader.Create(rStream);
            while (rReader.Read())
            {
                if (rReader.NodeType == XmlNodeType.Element && rReader.LocalName == "Relationship")
                {
                    string? rId = rReader.GetAttribute("Id");
                    string? target = rReader.GetAttribute("Target");
                    if (!string.IsNullOrEmpty(rId) && !string.IsNullOrEmpty(target))
                    {
                        relsMap[rId] = target;
                    }
                }
            }
        }

        // Parse sheets
        using var stream = entry.Open();
        using var reader = XmlReader.Create(stream);
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "sheet")
            {
                string? name = reader.GetAttribute("name");
                string? rId = reader.GetAttribute("r:id") ?? reader.GetAttribute("id"); // Handle namespace variations

                if (!string.IsNullOrEmpty(name) && !string.IsNullOrEmpty(rId) && relsMap.TryGetValue(rId, out string? relativePath))
                {
                    string fullPath = relativePath.StartsWith("/") ? relativePath.Substring(1) : $"xl/{relativePath}";
                    _sheetMap[name] = fullPath;
                }
            }
        }
    }

    private object? ParseCellValue(string raw, string? type)
    {
        if (string.IsNullOrEmpty(raw)) return null;

        if (type == "s" && int.TryParse(raw, out int strIndex))
        {
            return strIndex >= 0 && strIndex < _sharedStrings.Count ? _sharedStrings[strIndex] : string.Empty;
        }

        if (type == "b")
        {
            return raw == "1" || string.Equals(raw, "true", StringComparison.OrdinalIgnoreCase);
        }

        // Numeric or general value
        if (double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double d))
        {
            return d;
        }

        return raw;
    }

    private static string ReadInlineString(XmlReader reader)
    {
        while (reader.Read())
        {
            if (reader.NodeType == XmlNodeType.Element && reader.LocalName == "t")
            {
                return reader.ReadElementContentAsString();
            }
        }
        return string.Empty;
    }

    private static int GetColumnIndex(string address)
    {
        int col = 0;
        for (int i = 0; i < address.Length; i++)
        {
            char c = address[i];
            if (char.IsDigit(c)) break;
            if (c >= 'A' && c <= 'Z')
            {
                col = col * 26 + (c - 'A' + 1);
            }
            else if (c >= 'a' && c <= 'z')
            {
                col = col * 26 + (c - 'a' + 1);
            }
        }
        return col;
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (!_isDisposed)
        {
            _archive.Dispose();
            _isDisposed = true;
        }
    }
}
