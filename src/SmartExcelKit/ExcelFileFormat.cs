namespace SmartExcelKit;

/// <summary>
/// Supported spreadsheet and text-delimited formats in SmartExcelKit.
/// </summary>
public enum ExcelFileFormat
{
    /// <summary>
    /// Unknown or format not yet detected.
    /// </summary>
    Unknown,

    /// <summary>
    /// Excel Open XML Spreadsheet (.xlsx)
    /// </summary>
    Xlsx,

    /// <summary>
    /// Excel Open XML Macro-Enabled Spreadsheet (.xlsm)
    /// </summary>
    Xlsm,

    /// <summary>
    /// Excel Binary Spreadsheet (.xlsb)
    /// </summary>
    Xlsb,

    /// <summary>
    /// Legacy Excel 97-2003 binary format (.xls)
    /// </summary>
    Xls,

    /// <summary>
    /// Comma-Separated Values (.csv)
    /// </summary>
    Csv,

    /// <summary>
    /// Tab-Separated Values (.tsv)
    /// </summary>
    Tsv,

    /// <summary>
    /// Text file, custom delimiter delimited.
    /// </summary>
    Txt,

    /// <summary>
    /// XML Spreadsheet 2003 schema format.
    /// </summary>
    Xml2003,

    /// <summary>
    /// HTML table parsing/export.
    /// </summary>
    HtmlTable,

    /// <summary>
    /// JSON serialization format.
    /// </summary>
    Json
}
