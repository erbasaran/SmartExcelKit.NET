using System.Buffers;
using System.Text;

namespace SmartExcelKit.Csv;

/// <summary>
/// A high-performance, low-allocation CSV engine for reading and writing large delimited files.
/// </summary>
public static class CsvEngine
{
    private static readonly char[] DefaultDelimiters = [',', ';', '\t', '|'];

    /// <summary>
    /// Detects encoding from a seekable stream by checking byte order marks (BOM).
    /// </summary>
    public static Encoding DetectEncoding(Stream stream)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanSeek) return Encoding.UTF8;

        long originalPosition = stream.Position;
        byte[] bom = new byte[4];
        int read = stream.Read(bom, 0, bom.Length);
        stream.Position = originalPosition; // Reset positions

        if (read >= 3 && bom[0] == 0xEF && bom[1] == 0xBB && bom[2] == 0xBF)
            return Encoding.UTF8;
        if (read >= 2 && bom[0] == 0xFF && bom[1] == 0xFE)
            return Encoding.Unicode; // UTF-16 LE
        if (read >= 2 && bom[0] == 0xFE && bom[1] == 0xFF)
            return Encoding.BigEndianUnicode; // UTF-16 BE

        // Try registering code pages for Turkish / ISO-8859-9 / Windows-1254 support if running on .NET Core
        try
        {
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        }
        catch
        {
            // Fallback if provider is not available
        }

        return Encoding.UTF8;
    }

    /// <summary>
    /// Detects the delimiter by counting occurrences of potential separators in the first line.
    /// </summary>
    public static char DetectDelimiter(Stream stream, Encoding encoding)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanSeek) return ',';

        long originalPosition = stream.Position;
        using (var reader = new StreamReader(stream, encoding, detectEncodingFromByteOrderMarks: true, bufferSize: 1024, leaveOpen: true))
        {
            string? firstLine = reader.ReadLine();
            stream.Position = originalPosition; // Reset position

            if (string.IsNullOrEmpty(firstLine)) return ',';

            char bestDelimiter = ',';
            int maxCount = -1;

            foreach (char delimiter in DefaultDelimiters)
            {
                int count = 0;
                bool inQuotes = false;
                for (int i = 0; i < firstLine.Length; i++)
                {
                    char c = firstLine[i];
                    if (c == '"')
                    {
                        inQuotes = !inQuotes;
                    }
                    else if (c == delimiter && !inQuotes)
                    {
                        count++;
                    }
                }

                if (count > maxCount)
                {
                    maxCount = count;
                    bestDelimiter = delimiter;
                }
            }

            return maxCount > 0 ? bestDelimiter : ',';
        }
    }

    /// <summary>
    /// Reads and parses all CSV rows from a stream.
    /// </summary>
    public static List<List<string>> Read(Stream stream, char? delimiter = null, Encoding? encoding = null)
    {
        var result = new List<List<string>>();
        foreach (var row in ReadStreaming(stream, delimiter, encoding))
        {
            result.Add(row);
        }
        return result;
    }

    /// <summary>
    /// Parses a CSV stream row-by-row, supporting extremely large files without holding them in memory.
    /// </summary>
    public static IEnumerable<List<string>> ReadStreaming(Stream stream, char? delimiter = null, Encoding? encoding = null)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));

        var enc = encoding ?? DetectEncoding(stream);
        char delim = delimiter ?? DetectDelimiter(stream, enc);

        using var reader = new StreamReader(stream, enc, detectEncodingFromByteOrderMarks: true, bufferSize: 4096, leaveOpen: true);

        var rowFields = new List<string>();
        var fieldBuilder = new StringBuilder();

        const int bufferSize = 4096;
        char[] buffer = ArrayPool<char>.Shared.Rent(bufferSize);

        try
        {
            bool inQuotes = false;
            bool wasQuote = false;
            int readCount;

            while ((readCount = reader.Read(buffer, 0, bufferSize)) > 0)
            {
                for (int i = 0; i < readCount; i++)
                {
                    char c = buffer[i];

                    if (inQuotes)
                    {
                        if (c == '"')
                        {
                            inQuotes = false;
                            wasQuote = true;
                        }
                        else
                        {
                            fieldBuilder.Append(c);
                        }
                    }
                    else
                    {
                        if (wasQuote)
                        {
                            wasQuote = false;
                            if (c == '"')
                            {
                                // Escaped quote: "" -> "
                                fieldBuilder.Append('"');
                                inQuotes = true;
                                continue;
                            }
                        }

                        if (c == '"')
                        {
                            inQuotes = true;
                        }
                        else if (c == delim)
                        {
                            rowFields.Add(fieldBuilder.Length == 0 ? string.Empty : fieldBuilder.ToString());
                            fieldBuilder.Clear();
                        }
                        else if (c == '\r')
                        {
                            // Peek to see if \n follows
                            if (i + 1 < readCount && buffer[i + 1] == '\n')
                            {
                                i++; // Consume \n
                            }
                            rowFields.Add(fieldBuilder.Length == 0 ? string.Empty : fieldBuilder.ToString());
                            fieldBuilder.Clear();
                            yield return rowFields;
                            rowFields = new List<string>();
                        }
                        else if (c == '\n')
                        {
                            rowFields.Add(fieldBuilder.Length == 0 ? string.Empty : fieldBuilder.ToString());
                            fieldBuilder.Clear();
                            yield return rowFields;
                            rowFields = new List<string>();
                        }
                        else
                        {
                            fieldBuilder.Append(c);
                        }
                    }
                }
            }

            // Yield trailing fields/rows
            if (fieldBuilder.Length > 0 || rowFields.Count > 0)
            {
                rowFields.Add(fieldBuilder.Length == 0 ? string.Empty : fieldBuilder.ToString());
                yield return rowFields;
            }
        }
        finally
        {
            ArrayPool<char>.Shared.Return(buffer);
        }
    }

    /// <summary>
    /// Writes all rows to a CSV stream.
    /// </summary>
    public static void Write(Stream stream, IEnumerable<List<string>> rows, char delimiter = ',', Encoding? encoding = null)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (rows == null) throw new ArgumentNullException(nameof(rows));

        var enc = encoding ?? Encoding.UTF8;
        using var writer = new StreamWriter(stream, enc, bufferSize: 4096, leaveOpen: true);

        foreach (var row in rows)
        {
            WriteRow(writer, row, delimiter);
        }
        writer.Flush();
    }

    /// <summary>
    /// Writes a single CSV row to a StreamWriter, escaping quotes and delimiters.
    /// </summary>
    private static void WriteRow(StreamWriter writer, List<string> fields, char delimiter)
    {
        for (int i = 0; i < fields.Count; i++)
        {
            if (i > 0) writer.Write(delimiter);

            string val = fields[i] ?? string.Empty;
            bool needsQuotes = val.IndexOf(delimiter) >= 0 || val.IndexOf('"') >= 0 || val.IndexOf('\n') >= 0 || val.IndexOf('\r') >= 0;

            if (needsQuotes)
            {
                writer.Write('"');
                writer.Write(val.Replace("\"", "\"\""));
                writer.Write('"');
            }
            else
            {
                writer.Write(val);
            }
        }
        writer.WriteLine();
    }
}
