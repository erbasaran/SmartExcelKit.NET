using SmartExcelKit.Exceptions;
using System.Text;

namespace SmartExcelKit.Providers;

/// <summary>
/// Format provider to import and export workbooks in HTML table format.
/// </summary>
public sealed class HtmlTableFormatProvider : IWorkbookFormatProvider
{
    /// <inheritdoc />
    public void Read(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            using var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 2048, leaveOpen: true);
            string html = reader.ReadToEnd();

            // Clear existing sheets
            while (workbook.Worksheets.Count > 0)
            {
                workbook.RemoveWorksheet(workbook.Worksheets[0].Name);
            }

            var sheet = workbook.AddWorksheet("Sheet1");
            int rowIndex = 1;

            int trStart = 0;
            while ((trStart = html.IndexOf("<tr", trStart, StringComparison.OrdinalIgnoreCase)) != -1)
            {
                int trClose = html.IndexOf(">", trStart);
                if (trClose == -1) break;

                int trEnd = html.IndexOf("</tr>", trClose, StringComparison.OrdinalIgnoreCase);
                if (trEnd == -1) break;

                string trContent = html.Substring(trClose + 1, trEnd - trClose - 1);
                int colIndex = 1;

                int tdStart = 0;
                while (true)
                {
                    int tdOpen = trContent.IndexOf("<td", tdStart, StringComparison.OrdinalIgnoreCase);
                    int thOpen = trContent.IndexOf("<th", tdStart, StringComparison.OrdinalIgnoreCase);

                    int currentTdStart;
                    string closeTag;
                    if (tdOpen != -1 && thOpen != -1)
                    {
                        if (tdOpen < thOpen) { currentTdStart = tdOpen; closeTag = "</td>"; }
                        else { currentTdStart = thOpen; closeTag = "</th>"; }
                    }
                    else if (tdOpen != -1) { currentTdStart = tdOpen; closeTag = "</td>"; }
                    else if (thOpen != -1) { currentTdStart = thOpen; closeTag = "</th>"; }
                    else break;

                    int tdClose = trContent.IndexOf(">", currentTdStart);
                    if (tdClose == -1) break;

                    int tdEnd = trContent.IndexOf(closeTag, tdClose, StringComparison.OrdinalIgnoreCase);
                    if (tdEnd == -1) break;

                    string cellText = trContent.Substring(tdClose + 1, tdEnd - tdClose - 1).Trim();

                    // Decode basic HTML entities
                    cellText = cellText.Replace("&nbsp;", " ")
                                       .Replace("&amp;", "&")
                                       .Replace("&lt;", "<")
                                       .Replace("&gt;", ">")
                                       .Replace("&quot;", "\"");

                    if (double.TryParse(cellText, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double d))
                        sheet.Cell(rowIndex, colIndex).Value = d;
                    else if (bool.TryParse(cellText, out bool b))
                        sheet.Cell(rowIndex, colIndex).Value = b;
                    else
                        sheet.Cell(rowIndex, colIndex).Value = cellText;

                    colIndex++;
                    tdStart = tdEnd + closeTag.Length;
                }

                if (colIndex > 1)
                {
                    rowIndex++;
                }
                trStart = trEnd + 5;
            }
        }
        catch (Exception ex)
        {
            throw new ParsingException("Failed to parse HTML tables.", "HTML_PARSING_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public void Write(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        if (workbook.Worksheets.Count == 0)
            throw new ExportException("Workbook must contain at least one worksheet to export to HTML.", "NO_WORKSHEET_TO_EXPORT");

        var sheet = workbook.Worksheets[0];
        int maxRow = sheet.MaxRow;
        int maxCol = sheet.MaxColumn;

        using var writer = new StreamWriter(stream, Encoding.UTF8, bufferSize: 2048, leaveOpen: true);

        writer.WriteLine("<!DOCTYPE html>");
        writer.WriteLine("<html>");
        writer.WriteLine("<head>");
        writer.WriteLine("<meta charset=\"utf-8\" />");
        writer.WriteLine("<style>");
        writer.WriteLine("  table { border-collapse: collapse; width: 100%; font-family: sans-serif; }");
        writer.WriteLine("  td, th { border: 1px solid #dddddd; text-align: left; padding: 8px; }");
        writer.WriteLine("  tr:nth-child(even) { background-color: #f2f2f2; }");
        writer.WriteLine("</style>");
        writer.WriteLine("</head>");
        writer.WriteLine("<body>");
        writer.WriteLine("<table>");

        for (int r = 1; r <= maxRow; r++)
        {
            writer.Write("  <tr>");
            for (int c = 1; c <= maxCol; c++)
            {
                string rawVal = sheet.Cell(r, c).GetString();
                // Encode HTML special chars
                string encoded = rawVal.Replace("&", "&amp;")
                                       .Replace("<", "&lt;")
                                       .Replace(">", "&gt;")
                                       .Replace("\"", "&quot;");

                writer.Write($"<td>{encoded}</td>");
            }
            writer.WriteLine("</tr>");
        }

        writer.WriteLine("</table>");
        writer.WriteLine("</body>");
        writer.WriteLine("</html>");
        writer.Flush();
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
}
