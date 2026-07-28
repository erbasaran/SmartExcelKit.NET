using System.Globalization;
using System.Text;

namespace SmartExcelKit.Csv;

/// <summary>
/// Internal utility providing zero-allocation and safe CSV string formatting.
/// </summary>
internal static class CsvHelper
{
    /// <summary>
    /// Formats an array of cell values into a CSV formatted line string.
    /// </summary>
    public static string FormatRow(object?[] values, char delimiter = ',', char quoteChar = '"', CultureInfo? culture = null)
    {
        if (values == null || values.Length == 0) return string.Empty;

        culture ??= CultureInfo.InvariantCulture;
        var sb = new StringBuilder();

        for (int i = 0; i < values.Length; i++)
        {
            if (i > 0) sb.Append(delimiter);

            var val = values[i];
            if (val == null || val is DBNull) continue;

            string str = val is IFormattable formattable
                ? formattable.ToString(null, culture)
                : val.ToString() ?? string.Empty;

            bool needsQuotes = str.IndexOf(delimiter) >= 0 || str.IndexOf(quoteChar) >= 0 || str.IndexOf('\n') >= 0 || str.IndexOf('\r') >= 0;
            if (needsQuotes)
            {
                sb.Append(quoteChar);
                for (int j = 0; j < str.Length; j++)
                {
                    char c = str[j];
                    if (c == quoteChar) sb.Append(quoteChar);
                    sb.Append(c);
                }
                sb.Append(quoteChar);
            }
            else
            {
                sb.Append(str);
            }
        }

        return sb.ToString();
    }

    /// <summary>
    /// Formats a single sequence of cell values into a CSV formatted line string.
    /// </summary>
    public static string FormatRow(IEnumerable<object?> values, char delimiter = ',', char quoteChar = '"', CultureInfo? culture = null)
    {
        if (values == null) return string.Empty;
        if (values is object?[] arr) return FormatRow(arr, delimiter, quoteChar, culture);

        culture ??= CultureInfo.InvariantCulture;
        var sb = new StringBuilder();
        bool first = true;

        foreach (var val in values)
        {
            if (!first) sb.Append(delimiter);
            first = false;

            if (val == null || val is DBNull) continue;

            string str = val is IFormattable formattable
                ? formattable.ToString(null, culture)
                : val.ToString() ?? string.Empty;

            bool needsQuotes = str.IndexOf(delimiter) >= 0 || str.IndexOf(quoteChar) >= 0 || str.IndexOf('\n') >= 0 || str.IndexOf('\r') >= 0;
            if (needsQuotes)
            {
                sb.Append(quoteChar);
                for (int i = 0; i < str.Length; i++)
                {
                    char c = str[i];
                    if (c == quoteChar) sb.Append(quoteChar);
                    sb.Append(c);
                }
                sb.Append(quoteChar);
            }
            else
            {
                sb.Append(str);
            }
        }

        return sb.ToString();
    }
}
