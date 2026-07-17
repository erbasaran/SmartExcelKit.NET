using SmartExcelKit.Exceptions;
using System.Text.Json;

namespace SmartExcelKit.Providers;

/// <summary>
/// Format provider to import and export workbooks in JSON format.
/// </summary>
public sealed class JsonFormatProvider : IWorkbookFormatProvider
{
    /// <inheritdoc />
    public void Read(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            using var reader = new StreamReader(stream, System.Text.Encoding.UTF8, detectEncodingFromByteOrderMarks: true, bufferSize: 2048, leaveOpen: true);
            string json = reader.ReadToEnd();

            var data = JsonSerializer.Deserialize<Dictionary<string, List<List<object>>>>(json);
            if (data == null) return;

            // Clear existing sheets
            while (workbook.Worksheets.Count > 0)
            {
                workbook.RemoveWorksheet(workbook.Worksheets[0].Name);
            }

            foreach (var kvp in data)
            {
                var sheet = workbook.AddWorksheet(kvp.Key);
                for (int r = 0; r < kvp.Value.Count; r++)
                {
                    var rowData = kvp.Value[r];
                    for (int c = 0; c < rowData.Count; c++)
                    {
                        var cellVal = rowData[c];
                        if (cellVal is JsonElement elem)
                        {
                            sheet.Cell(r + 1, c + 1).Value = ExtractJsonValue(elem);
                        }
                        else
                        {
                            sheet.Cell(r + 1, c + 1).Value = cellVal;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            throw new SerializationException("Failed to deserialize Excel JSON content.", "JSON_DESERIALIZATION_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public void Write(Stream stream, ExcelWorkbook workbook)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (workbook == null) throw new ArgumentNullException(nameof(workbook));

        try
        {
            var data = new Dictionary<string, List<List<object?>>>();
            foreach (var sheet in workbook.Worksheets)
            {
                var rows = new List<List<object?>>();
                int maxRow = sheet.MaxRow;
                int maxCol = sheet.MaxColumn;

                for (int r = 1; r <= maxRow; r++)
                {
                    var rowFields = new List<object?>();
                    for (int c = 1; c <= maxCol; c++)
                    {
                        rowFields.Add(sheet.Cell(r, c).Value);
                    }
                    rows.Add(rowFields);
                }
                data[sheet.Name] = rows;
            }

            var options = new JsonSerializerOptions { WriteIndented = true };
            var jsonBytes = JsonSerializer.SerializeToUtf8Bytes(data, options);
            stream.Write(jsonBytes, 0, jsonBytes.Length);
        }
        catch (Exception ex)
        {
            throw new SerializationException("Failed to serialize Excel to JSON.", "JSON_SERIALIZATION_FAILED", ex);
        }
    }

    /// <inheritdoc />
    public async Task ReadAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        using var reader = new StreamReader(stream, System.Text.Encoding.UTF8, true, 2048, true);
        string json = await reader.ReadToEndAsync();

        var data = JsonSerializer.Deserialize<Dictionary<string, List<List<object>>>>(json);
        if (data == null) return;

        while (workbook.Worksheets.Count > 0)
        {
            workbook.RemoveWorksheet(workbook.Worksheets[0].Name);
        }

        foreach (var kvp in data)
        {
            var sheet = workbook.AddWorksheet(kvp.Key);
            for (int r = 0; r < kvp.Value.Count; r++)
            {
                var rowData = kvp.Value[r];
                for (int c = 0; c < rowData.Count; c++)
                {
                    var cellVal = rowData[c];
                    if (cellVal is JsonElement elem)
                    {
                        sheet.Cell(r + 1, c + 1).Value = ExtractJsonValue(elem);
                    }
                    else
                    {
                        sheet.Cell(r + 1, c + 1).Value = cellVal;
                    }
                }
            }
        }
    }

    /// <inheritdoc />
    public async Task WriteAsync(Stream stream, ExcelWorkbook workbook, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var data = new Dictionary<string, List<List<object?>>>();
        foreach (var sheet in workbook.Worksheets)
        {
            var rows = new List<List<object?>>();
            int maxRow = sheet.MaxRow;
            int maxCol = sheet.MaxColumn;

            for (int r = 1; r <= maxRow; r++)
            {
                var rowFields = new List<object?>();
                for (int c = 1; c <= maxCol; c++)
                {
                    rowFields.Add(sheet.Cell(r, c).Value);
                }
                rows.Add(rowFields);
            }
            data[sheet.Name] = rows;
        }

        var options = new JsonSerializerOptions { WriteIndented = true };
#if NETSTANDARD2_0
        var jsonBytes = JsonSerializer.SerializeToUtf8Bytes(data, options);
        await stream.WriteAsync(jsonBytes, 0, jsonBytes.Length, cancellationToken);
#else
        await JsonSerializer.SerializeAsync(stream, data, options, cancellationToken);
#endif
    }

    private static object? ExtractJsonValue(JsonElement element)
    {
        return element.ValueKind switch
        {
            JsonValueKind.String => element.GetString(),
            JsonValueKind.Number => element.TryGetDouble(out double d) ? d : element.GetRawText(),
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.Null => null,
            _ => element.GetRawText()
        };
    }
}
