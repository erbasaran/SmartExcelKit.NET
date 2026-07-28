using System.Data;
using System.Text;
using SmartExcelKit.Core;
using SmartExcelKit.Exceptions;
using SmartExcelKit.Providers;
using SmartExcelKit.Styles;

namespace SmartExcelKit;

/// <summary>
/// Represents an Excel Workbook containing worksheets, global styles, and document properties.
/// </summary>
public sealed class ExcelWorkbook : IDisposable
{
    private readonly List<ExcelWorksheet> _worksheets = [];
    private readonly StyleRegistry _styleRegistry = new();
    private readonly DocumentProperties _properties = new();

    private int _activeWorksheetIndex = 0;

    /// <summary>
    /// Gets the list of worksheets in the workbook.
    /// </summary>
    public IReadOnlyList<ExcelWorksheet> Worksheets => _worksheets;

    /// <summary>
    /// Gets or sets the active worksheet. Returns the first worksheet if not set, or creates a default worksheet if none exists.
    /// </summary>
    public ExcelWorksheet ActiveWorksheet
    {
        get
        {
            if (_worksheets.Count == 0)
            {
                return AddWorksheet("Sheet1");
            }
            if (_activeWorksheetIndex < 0 || _activeWorksheetIndex >= _worksheets.Count)
            {
                _activeWorksheetIndex = 0;
            }
            return _worksheets[_activeWorksheetIndex];
        }
        set
        {
            if (value == null) throw new ArgumentNullException(nameof(value));
            int idx = _worksheets.IndexOf(value);
            if (idx < 0)
                throw new ArgumentException("The specified worksheet does not belong to this workbook.", nameof(value));
            _activeWorksheetIndex = idx;
        }
    }

    /// <summary>
    /// Gets a worksheet by name (case-insensitive).
    /// </summary>
    /// <param name="name">The worksheet name.</param>
    /// <returns>The matching <see cref="ExcelWorksheet"/>.</returns>
    /// <exception cref="WorksheetException">Thrown if no worksheet with the specified name is found.</exception>
    public ExcelWorksheet this[string name]
    {
        get
        {
            if (string.IsNullOrWhiteSpace(name))
                throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(name));

            var ws = _worksheets.Find(w => string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase));
            if (ws == null)
            {
                throw new WorksheetException($"Worksheet '{name}' was not found in the workbook.", "WORKSHEET_NOT_FOUND");
            }
            return ws;
        }
    }

    /// <summary>
    /// Gets a worksheet by 0-based index.
    /// </summary>
    /// <param name="index">The 0-based worksheet index.</param>
    /// <returns>The <see cref="ExcelWorksheet"/> at the specified index.</returns>
    public ExcelWorksheet this[int index] => _worksheets[index];

    /// <summary>
    /// Gets the global style registry.
    /// </summary>
    public StyleRegistry StyleRegistry => _styleRegistry;

    /// <summary>
    /// Gets the document properties and metadata.
    /// </summary>
    public DocumentProperties Properties => _properties;

    /// <summary>
    /// Gets the collection of workbook-scoped named ranges.
    /// </summary>
    public Core.ExcelNamedRangeCollection NamedRanges { get; } = new Core.ExcelNamedRangeCollection();

    /// <summary>
    /// Gets or sets raw VBA binary project bytes (vbaProject.bin) to preserve macros in macro-enabled templates/files (.xlsm).
    /// </summary>
    public byte[]? VbaProjectBytes { get; set; }

    /// <summary>
    /// Gets the workbook protection password hash (HEX format). Returns null if workbook is not protected.
    /// </summary>
    public string? ProtectionPasswordHash { get; internal set; }

    /// <summary>
    /// Gets a value indicating whether the workbook structure is protected.
    /// </summary>
    public bool IsProtected => ProtectionPasswordHash != null;

    /// <summary>
    /// Protects the workbook structure with a password using the Excel XOR hashing algorithm.
    /// </summary>
    public void Protect(string password)
    {
        if (string.IsNullOrEmpty(password))
        {
            ProtectionPasswordHash = "0000";
            return;
        }

        int hash = 0;
        for (int i = password.Length - 1; i >= 0; i--)
        {
            char c = password[i];
            hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
            hash ^= c;
        }
        hash = ((hash >> 14) & 0x01) | ((hash << 1) & 0x7FFF);
        hash ^= password.Length;
        hash ^= 0xCE4B;

        ProtectionPasswordHash = hash.ToString("X4");
    }

    /// <summary>
    /// Unprotects the workbook.
    /// </summary>
    public void Unprotect()
    {
        ProtectionPasswordHash = null;
    }

    /// <summary>
    /// Recalculates all cell formulas across all worksheets in the workbook.
    /// </summary>
    public void Recalculate()
    {
        Formula.FormulaEvaluator.ClearCache();
        foreach (var ws in _worksheets)
        {
            foreach (var cell in ws.CellsUsed())
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    cell.Value = Formula.FormulaEvaluator.Evaluate(cell.Formula!, ws, cell.Address);
                }
            }
        }
    }

    /// <summary>
    /// Initializes a new, empty <see cref="ExcelWorkbook"/>.
    /// </summary>
    public ExcelWorkbook()
    {
    }

    /// <summary>
    /// Creates a new worksheet in this workbook with a unique name.
    /// </summary>
    /// <param name="name">The name of the worksheet.</param>
    /// <returns>A new <see cref="ExcelWorksheet"/> instance.</returns>
    public ExcelWorksheet AddWorksheet(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Worksheet name cannot be null or empty.", nameof(name));

        foreach (var ws in _worksheets)
        {
            if (string.Equals(ws.Name, name, StringComparison.OrdinalIgnoreCase))
                throw new WorksheetException($"A worksheet with the name '{name}' already exists.", "DUPLICATE_WORKSHEET_NAME");
        }

        var worksheet = new ExcelWorksheet(this, name);
        _worksheets.Add(worksheet);
        return worksheet;
    }

    /// <summary>
    /// Removes a worksheet from the workbook.
    /// </summary>
    /// <param name="name">The name of the worksheet to remove.</param>
    public void RemoveWorksheet(string name)
    {
        var ws = _worksheets.Find(w => string.Equals(w.Name, name, StringComparison.OrdinalIgnoreCase));
        if (ws != null)
        {
            _worksheets.Remove(ws);
        }
    }

    #region Save Methods

    /// <summary>
    /// Saves the workbook to the specified file path, auto-detecting the format.
    /// </summary>
    public void Save(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Path cannot be null or empty.", nameof(path));

        var format = DetectFormatFromExtension(path);
        using var stream = File.Create(path);
        Save(stream, format);
    }

    /// <summary>
    /// Saves the workbook to the specified stream with the given format.
    /// </summary>
    public void Save(Stream stream, ExcelFileFormat format)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (format == ExcelFileFormat.Unknown)
            throw new ArgumentException("Must specify a valid file format for streaming output.", nameof(format));

        var provider = FormatProviderRegistry.Resolve(format);
        provider.Write(stream, this);
    }

    /// <summary>
    /// Saves the workbook asynchronously to the specified file path.
    /// </summary>
    public async Task SaveAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Path cannot be null or empty.", nameof(path));

        var format = DetectFormatFromExtension(path);
        using var stream = File.Create(path);
        await SaveAsync(stream, format, cancellationToken);
    }

    /// <summary>
    /// Saves the workbook asynchronously to the specified stream.
    /// </summary>
    public async Task SaveAsync(Stream stream, ExcelFileFormat format, CancellationToken cancellationToken = default)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (format == ExcelFileFormat.Unknown)
            throw new ArgumentException("Must specify a valid file format for streaming output.", nameof(format));

        var provider = FormatProviderRegistry.Resolve(format);
        await provider.WriteAsync(stream, this, cancellationToken);
    }

    #endregion

    #region Open Methods

    /// <summary>
    /// Opens a workbook from the specified file path, auto-detecting the format.
    /// </summary>
    public static ExcelWorkbook Open(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Path cannot be null or empty.", nameof(path));

        using var stream = File.OpenRead(path);
        var format = DetectFormat(stream, path);
        var workbook = new ExcelWorkbook();
        workbook.Load(stream, format);
        return workbook;
    }

    /// <summary>
    /// Opens a workbook from a stream, auto-detecting format.
    /// </summary>
    public static ExcelWorkbook Open(Stream stream)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        var format = DetectFormat(stream);
        var workbook = new ExcelWorkbook();
        workbook.Load(stream, format);
        return workbook;
    }
    /// <summary>
    /// Opens an Excel workbook from a byte array.
    /// The file format is auto-detected.
    /// </summary>
    /// <param name="data">The raw workbook bytes.</param>
    /// <returns>A loaded <see cref="ExcelWorkbook"/>.</returns>
    public static ExcelWorkbook Open(byte[] data)
    {
        if (data == null) throw new ArgumentNullException(nameof(data));
        using var ms = new MemoryStream(data);
        return Open(ms);
    }
    /// <summary>
    /// Opens a workbook asynchronously from a file path.
    /// </summary>
    public static async Task<ExcelWorkbook> OpenAsync(string path, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(path))
            throw new ArgumentException("Path cannot be null or empty.", nameof(path));

        using var stream = File.OpenRead(path);
        var format = DetectFormat(stream, path);
        var workbook = new ExcelWorkbook();
        await workbook.LoadAsync(stream, format, cancellationToken);
        return workbook;
    }

    /// <summary>
    /// Opens a workbook asynchronously from a stream.
    /// </summary>
    public static async Task<ExcelWorkbook> OpenAsync(Stream stream, CancellationToken cancellationToken = default)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        var format = DetectFormat(stream);
        var workbook = new ExcelWorkbook();
        await workbook.LoadAsync(stream, format, cancellationToken);
        return workbook;
    }

    private void Load(Stream stream, ExcelFileFormat format)
    {
        var provider = FormatProviderRegistry.Resolve(format);
        provider.Read(stream, this);
    }

    private async Task LoadAsync(Stream stream, ExcelFileFormat format, CancellationToken cancellationToken)
    {
        var provider = FormatProviderRegistry.Resolve(format);
        await provider.ReadAsync(stream, this, cancellationToken);
    }

    #endregion

    #region Format Detection Utilities

    /// <summary>
    /// Detects the file format by checking the stream contents and optional file path.
    /// </summary>
    public static ExcelFileFormat DetectFormat(Stream stream, string? path = null)
    {
        if (stream == null) throw new ArgumentNullException(nameof(stream));
        if (!stream.CanSeek)
        {
            // If stream is forward-only, we must fall back to file path detection
            if (!string.IsNullOrEmpty(path))
            {
                return DetectFormatFromExtension(path!);
            }
            throw new ArgumentException("Cannot detect format from a non-seekable stream without a file path suffix.", nameof(stream));
        }

        long initialPos = stream.Position;
        byte[] buffer = new byte[128];
        int readBytes = stream.Read(buffer, 0, buffer.Length);
        stream.Position = initialPos; // Restore stream position

        if (readBytes >= 8 &&
            buffer[0] == 0xD0 && buffer[1] == 0xCF && buffer[2] == 0x11 && buffer[3] == 0xE0 &&
            buffer[4] == 0xA1 && buffer[5] == 0xB1 && buffer[6] == 0x1A && buffer[7] == 0xE1)
        {
            return ExcelFileFormat.Xls;
        }

        if (readBytes >= 4)
        {
            // ZIP magic number: PK\x03\x04 (50 4B 03 04)
            if (buffer[0] == 0x50 && buffer[1] == 0x4B && buffer[2] == 0x03 && buffer[3] == 0x04)
            {
                if (!string.IsNullOrEmpty(path))
                {
                    var ext = Path.GetExtension(path!).ToUpperInvariant();
                    if (ext == ".XLSM") return ExcelFileFormat.Xlsm;
                    if (ext == ".XLSB") return ExcelFileFormat.Xlsb;
                }
                return ExcelFileFormat.Xlsx;
            }

            // Strip BOM bytes if present
            int offset = 0;
            if (readBytes >= 3 && buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
            {
                offset = 3;
            }
            else if (readBytes >= 2 && buffer[0] == 0xFF && buffer[1] == 0xFE)
            {
                offset = 2;
            }
            else if (readBytes >= 2 && buffer[0] == 0xFE && buffer[1] == 0xFF)
            {
                offset = 2;
            }

            string head = Encoding.UTF8.GetString(buffer, offset, readBytes - offset).TrimStart().ToUpperInvariant();

            // XML Spreadsheet 2003 detection
            if (head.StartsWith("<?XML") || head.StartsWith("<WORKBOOK") || head.Contains("<WORKBOOK") || head.Contains("<SS:WORKBOOK"))
            {
                return ExcelFileFormat.Xml2003;
            }

            // HTML Table detection
            if (head.StartsWith("<!DOCTYPE") || head.StartsWith("<HTML") || head.StartsWith("<TABLE") || head.Contains("<TABLE"))
            {
                return ExcelFileFormat.HtmlTable;
            }

            // JSON detection
            if (head.StartsWith("{") || head.StartsWith("["))
            {
                return ExcelFileFormat.Json;
            }
        }

        // Fallback to extension check if path is available
        if (!string.IsNullOrEmpty(path))
        {
            return DetectFormatFromExtension(path!);
        }

        return ExcelFileFormat.Csv; // Default fallback format
    }

    private static ExcelFileFormat DetectFormatFromExtension(string path)
    {
        var ext = Path.GetExtension(path).ToUpperInvariant();
        return ext switch
        {
            ".XLSX" => ExcelFileFormat.Xlsx,
            ".XLSM" => ExcelFileFormat.Xlsm,
            ".XLSB" => ExcelFileFormat.Xlsb,
            ".XLS" => ExcelFileFormat.Xls,
            ".CSV" => ExcelFileFormat.Csv,
            ".TSV" => ExcelFileFormat.Tsv,
            ".TXT" => ExcelFileFormat.Txt,
            ".JSON" => ExcelFileFormat.Json,
            ".XML" => ExcelFileFormat.Xml2003,
            ".HTML" or ".HTM" => ExcelFileFormat.HtmlTable,
            _ => ExcelFileFormat.Unknown
        };
    }

    #endregion

    #region Workbook Conversion & Export Helpers

    /// <summary>
    /// Formats all worksheets in the workbook into a single CSV formatted string.
    /// </summary>
    /// <param name="delimiter">The column delimiter character (default ',').</param>
    /// <param name="quoteCharacter">The quote character for escaping (default '"').</param>
    /// <returns>The combined CSV text across all worksheets.</returns>
    public string ToCsv(char delimiter = ',', char quoteCharacter = '"')
    {
        var sb = new StringBuilder();
        for (int i = 0; i < _worksheets.Count; i++)
        {
            var sheetCsv = _worksheets[i].ToCsv(delimiter, quoteCharacter);
            if (!string.IsNullOrWhiteSpace(sheetCsv))
            {
                if (sb.Length > 0)
                {
                    sb.AppendLine();
                }
                sb.Append(sheetCsv);
            }
        }
        return sb.ToString();
    }

    /// <summary>
    /// Serializes all worksheets in the workbook into a single JSON string.
    /// </summary>
    /// <param name="hasHeader">Whether top row contains property names (default true).</param>
    /// <returns>A JSON string representing all worksheets.</returns>
    public string ToJson(bool hasHeader = true)
    {
        var dict = new Dictionary<string, object>();
        foreach (var ws in _worksheets)
        {
            dict[ws.Name] = ws.ToDictionaryList(1, 1, hasHeader);
        }
        return System.Text.Json.JsonSerializer.Serialize(dict, new System.Text.Json.JsonSerializerOptions { WriteIndented = false });
    }

    /// <summary>
    /// Exports all worksheets in the workbook to a <see cref="DataSet"/>, where each table represents a worksheet.
    /// </summary>
    /// <param name="hasHeader">Whether top row contains column headers (default true).</param>
    /// <returns>A <see cref="DataSet"/> containing data tables for each worksheet.</returns>
    public DataSet ToDataSet(bool hasHeader = true)
    {
        var ds = new DataSet(Properties.Title ?? "WorkbookDataSet");
        foreach (var ws in _worksheets)
        {
            var dt = ws.ToDataTable(hasHeader);
            dt.TableName = ws.Name;
            ds.Tables.Add(dt);
        }
        return ds;
    }

    /// <summary>
    /// Exports active or combined worksheets to a single <see cref="DataTable"/>.
    /// </summary>
    /// <param name="hasHeader">Whether top row contains column headers (default true).</param>
    /// <returns>A <see cref="DataTable"/> representing worksheet data.</returns>
    public DataTable ToDataTable(bool hasHeader = true)
    {
        return ActiveWorksheet.ToDataTable(hasHeader);
    }

    /// <summary>
    /// Flattens and exports rows from ALL worksheets in the workbook to POCO objects of type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target POCO class type.</typeparam>
    /// <param name="startRow">The 1-based start row (default 1).</param>
    /// <param name="startColumn">The 1-based start column (default 1).</param>
    /// <returns>An enumerable sequence of POCO objects mapped across all worksheets.</returns>
    public IEnumerable<T> ToObjects<T>(int startRow = 1, int startColumn = 1) where T : class, new()
    {
        foreach (var ws in _worksheets)
        {
            foreach (var item in ws.ToObjects<T>(startRow, startColumn))
            {
                yield return item;
            }
        }
    }

    /// <summary>
    /// Converts all worksheets in the workbook to a list of row dictionaries.
    /// Includes a '__Worksheet' key indicating the originating sheet name.
    /// </summary>
    /// <param name="startRow">1-based start row index (default 1).</param>
    /// <param name="startColumn">1-based start column index (default 1).</param>
    /// <param name="hasHeader">Whether top row contains column headers (default true).</param>
    /// <returns>A list of dictionary rows representing data across all worksheets.</returns>
    public List<Dictionary<string, object?>> ToDictionaryList(int startRow = 1, int startColumn = 1, bool hasHeader = true)
    {
        var result = new List<Dictionary<string, object?>>();
        foreach (var ws in _worksheets)
        {
            var sheetRows = ws.ToDictionaryList(startRow, startColumn, hasHeader);
            foreach (var row in sheetRows)
            {
                if (!row.ContainsKey("__Worksheet"))
                {
                    row["__Worksheet"] = ws.Name;
                }
                result.Add(row);
            }
        }
        return result;
    }

    #endregion

    /// <inheritdoc />
    public void Dispose()
    {
        // Place for clean-up of temporary files or zip streams if added later
    }
}
