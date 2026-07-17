using SmartExcelKit.Exceptions;
using System.Collections.Concurrent;

namespace SmartExcelKit.Providers;

/// <summary>
/// Registry that resolves format providers based on ExcelFileFormat.
/// </summary>
public static class FormatProviderRegistry
{
    private static readonly ConcurrentDictionary<ExcelFileFormat, IWorkbookFormatProvider> _providers = new();

    static FormatProviderRegistry()
    {
        // Register default built-in providers
        Register(ExcelFileFormat.Csv, new CsvFormatProvider(delimiter: ','));
        Register(ExcelFileFormat.Tsv, new CsvFormatProvider(delimiter: '\t'));
    }

    /// <summary>
    /// Registers a provider for a specific file format.
    /// </summary>
    /// <param name="format">The file format.</param>
    /// <param name="provider">The provider instance.</param>
    public static void Register(ExcelFileFormat format, IWorkbookFormatProvider provider)
    {
        if (provider == null) throw new ArgumentNullException(nameof(provider));
        _providers[format] = provider;
    }

    /// <summary>
    /// Resolves the provider for a specific file format.
    /// </summary>
    /// <param name="format">The file format.</param>
    /// <returns>The resolved <see cref="IWorkbookFormatProvider"/>.</returns>
    /// <exception cref="WorkbookException">Thrown if no provider is registered for the specified format.</exception>
    public static IWorkbookFormatProvider Resolve(ExcelFileFormat format)
    {
        // Lazy-loading of complex format providers to avoid assembly load dependencies unless used
        if (!_providers.TryGetValue(format, out var provider))
        {
            provider = format switch
            {
                ExcelFileFormat.Xlsx or ExcelFileFormat.Xlsm => new XlsxFormatProvider(),
                ExcelFileFormat.Json => new JsonFormatProvider(),
                ExcelFileFormat.HtmlTable => new HtmlTableFormatProvider(),
                ExcelFileFormat.Xml2003 => new Xml2003FormatProvider(),
                ExcelFileFormat.Xls => new SmartExcelKit.Providers.Xls.XlsFormatProvider(),
                ExcelFileFormat.Xlsb =>
                    throw new WorkbookException($"Legacy binary format '{format}' is not supported directly in the core package. Register an external provider or convert files to XLSX.", "LEGACY_BINARY_FORMAT_UNSUPPORTED"),
                _ => throw new WorkbookException($"No format provider is registered for format: {format}", "PROVIDER_NOT_FOUND")
            };
            _providers[format] = provider;
        }
        return provider;
    }
}
