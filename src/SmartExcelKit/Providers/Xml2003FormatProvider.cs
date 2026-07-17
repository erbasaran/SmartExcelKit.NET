using SmartExcelKit.Exceptions;
using System.Xml.Linq;

namespace SmartExcelKit.Providers;

/// <summary>
/// Format provider to import and export XML Spreadsheet 2003 schema files.
/// </summary>
public sealed class Xml2003FormatProvider : IWorkbookFormatProvider
{
    private static readonly XNamespace ss = "urn:schemas-microsoft-com:office:spreadsheet";

    /// <inheritdoc />
    public void Read(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            var doc = XDocument.Load(stream);
            var root = doc.Root;
            if (root == null) throw new ParsingException("Invalid XML Spreadsheet 2003 document.", "XML2003_EMPTY");

            // Clear existing worksheets
            while (workbook.Worksheets.Count > 0)
            {
                workbook.RemoveWorksheet(workbook.Worksheets[0].Name);
            }

            var worksheets = root.Elements(ss + "Worksheet");
            foreach (var wsEl in worksheets)
            {
                string name = wsEl.Attribute(ss + "Name")?.Value ?? "Sheet";
                var sheet = workbook.AddWorksheet(name);

                var tableEl = wsEl.Element(ss + "Table");
                if (tableEl == null) continue;

                int rowIndex = 1;
                foreach (var rowEl in tableEl.Elements(ss + "Row"))
                {
                    var rowIdxAttr = rowEl.Attribute(ss + "Index");
                    if (rowIdxAttr != null)
                    {
                        rowIndex = int.Parse(rowIdxAttr.Value);
                    }

                    int colIndex = 1;
                    foreach (var cellEl in rowEl.Elements(ss + "Cell"))
                    {
                        var cellIdxAttr = cellEl.Attribute(ss + "Index");
                        if (cellIdxAttr != null)
                        {
                            colIndex = int.Parse(cellIdxAttr.Value);
                        }

                        var dataEl = cellEl.Element(ss + "Data");
                        if (dataEl != null)
                        {
                            string type = dataEl.Attribute(ss + "Type")?.Value ?? "String";
                            string rawVal = dataEl.Value;

                            sheet.Cell(rowIndex, colIndex).Value = ParseXmlDataValue(rawVal, type);
                        }

                        colIndex++;
                    }
                    rowIndex++;
                }
            }
        }
        catch (Exception ex) when (ex is not SmartExcelException)
        {
            throw new ParsingException("Failed to parse XML Spreadsheet 2003 document.", "XML2003_PARSE_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public void Write(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            var doc = new XDocument(new XDeclaration("1.0", "utf-8", "yes"));
            var root = new XElement(ss + "Workbook",
                new XAttribute(XNamespace.Xmlns + "ss", ss.NamespaceName),
                new XAttribute(XNamespace.Xmlns + "html", "http://www.w3.org/TR/REC-html40")
            );

            foreach (var sheet in workbook.Worksheets)
            {
                var wsEl = new XElement(ss + "Worksheet", new XAttribute(ss + "Name", sheet.Name));
                var tableEl = new XElement(ss + "Table");

                int maxRow = sheet.MaxRow;
                int maxCol = sheet.MaxColumn;

                for (int r = 1; r <= maxRow; r++)
                {
                    var rowEl = new XElement(ss + "Row");
                    bool rowHasData = false;

                    for (int c = 1; c <= maxCol; c++)
                    {
                        var val = sheet.Cell(r, c).Value;
                        if (val != null)
                        {
                            rowHasData = true;
                            var (typeStr, valStr) = FormatXmlDataValue(val);
                            var cellEl = new XElement(ss + "Cell",
                                new XAttribute(ss + "Index", c), // Write explicit column index
                                new XElement(ss + "Data",
                                    new XAttribute(ss + "Type", typeStr),
                                    valStr
                                )
                            );
                            rowEl.Add(cellEl);
                        }
                    }

                    if (rowHasData)
                    {
                        rowEl.Add(new XAttribute(ss + "Index", r)); // Write explicit row index
                        tableEl.Add(rowEl);
                    }
                }

                wsEl.Add(tableEl);
                root.Add(wsEl);
            }

            doc.Add(root);
            doc.Save(stream);
        }
        catch (Exception ex)
        {
            throw new ExportException("Failed to export XML Spreadsheet 2003.", "XML2003_EXPORT_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public Task ReadAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Read(stream, workbook), cancellationToken);
    }

    /// <inheritdoc />
    public Task WriteAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => Write(stream, workbook), cancellationToken);
    }

    private static object ParseXmlDataValue(string raw, string type)
    {
        return type.ToUpperInvariant() switch
        {
            "NUMBER" => double.TryParse(raw, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double d) ? d : raw,
            "BOOLEAN" => raw == "1" || string.Equals(raw, "TRUE", StringComparison.OrdinalIgnoreCase),
            "DATETIME" => DateTime.TryParse(raw, out DateTime dt) ? dt : raw,
            _ => raw
        };
    }

    private static (string Type, string Value) FormatXmlDataValue(object val)
    {
        if (val is double || val is float || val is int || val is decimal)
        {
            return ("Number", Convert.ToString(val, System.Globalization.CultureInfo.InvariantCulture));
        }
        if (val is bool b)
        {
            return ("Boolean", b ? "1" : "0");
        }
        if (val is DateTime dt)
        {
            return ("DateTime", dt.ToString("yyyy-MM-ddTHH:mm:ss.fff", System.Globalization.CultureInfo.InvariantCulture));
        }
        return ("String", val.ToString() ?? string.Empty);
    }
}
