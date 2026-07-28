# SmartExcelKit

[![NuGet Version](https://img.shields.io/nuget/v/SmartExcelKit.svg)](https://www.nuget.org/packages/SmartExcelKit)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

**SmartExcelKit** is a lightweight, ultra-fast, low-allocation .NET Standard 2.0 library for reading, writing, streaming, and evaluating spreadsheet files across multiple formats (**XLSX**, **XLS**, **CSV**, **TSV**, **HTML Table**, **XML Spreadsheet 2003**, **JSON**).

Designed with performance as the top priority, SmartExcelKit uses **pooled buffers (`ArrayPool<T>`)**, **lazy enumerations**, **O(1) bounds caching**, and **forward-only streaming** to deliver maximum throughput with minimal GC allocation.

---

## Target Frameworks

- **.NET Standard 2.0**: Compatible with .NET Core 2.0+, .NET 5+, .NET 6/7/8/9/10+, and .NET Framework 4.6.1+.

---

## Installation

```shell
dotnet add package SmartExcelKit
```

---

## Supported Formats

| Format | Extension | Read | Write | Description |
| :--- | :---: | :---: | :---: | :--- |
| **XLSX** | `.xlsx` | ✅ | ✅ | Standard Microsoft Excel OpenXML Workbook |
| **XLS** | `.xls` | ✅ | ❌ | Legacy Microsoft Excel Binary Format (BIFF8 / OLE2) |
| **CSV** | `.csv` | ✅ | ✅ | Comma-Separated Values |
| **TSV** | `.tsv` | ✅ | ✅ | Tab-Separated Values |
| **HTML** | `.html`, `.htm` | ✅ | ✅ | HTML Table Spreadsheet Markup |
| **XML 2003** | `.xml` | ✅ | ✅ | Microsoft Excel 2003 XML SpreadsheetML |
| **JSON** | `.json` | ✅ | ✅ | Structured JSON Array / Document |

---

## Developer Guide & Complete Code Examples

### 1. Open Workbooks (Auto-Format Detection)

Load workbooks from file paths, streams, or byte arrays. Format is auto-detected from magic byte signatures or extensions.

```csharp
using SmartExcelKit;
using SmartExcelKit.Core;

// 1. Open from file path (sync & async)
using var workbook1 = ExcelWorkbook.Open("data.xlsx");
using var workbook2 = await ExcelWorkbook.OpenAsync("data.csv");

// 2. Open from stream or byte array
using var stream = File.OpenRead("legacy.xls");
using var workbook3 = ExcelWorkbook.Open(stream);

byte[] bytes = File.ReadAllBytes("data.json");
using var workbook4 = ExcelWorkbook.Open(bytes);

// 3. Detect format explicitly
ExcelFileFormat format = ExcelWorkbook.DetectFormat(stream, "data.xlsx");
```

---

### 2. Create & Save Workbooks

Create workbooks, set active sheets, apply protection, and save to files or streams.

```csharp
using SmartExcelKit;

using var workbook = new ExcelWorkbook();
var sheet = workbook.AddWorksheet("MainSheet");

// Set active worksheet
workbook.ActiveWorksheet = sheet;

// Workbook protection
workbook.Protect("WorkbookPassword123");
bool isProtected = workbook.IsProtected;
workbook.Unprotect();

// Save to file or stream (sync & async)
workbook.Save("output.xlsx");
await workbook.SaveAsync("output.xlsx");

using var outStream = File.Create("output.csv");
workbook.Save(outStream, ExcelFileFormat.Csv);
```

---

### 3. Worksheet Management & Navigation

Access sheets by name or 0-based index, remove sheets, or protect individual worksheets.

```csharp
var sheet1 = workbook["MainSheet"];       // Access by name
var sheet2 = workbook[0];                 // Access by 0-based index

// Add / Remove worksheets
var newSheet = workbook.AddWorksheet("Draft");
workbook.RemoveWorksheet("Draft");

// Worksheet protection
sheet1.Protect("SheetPassword123");
sheet1.Unprotect();
```

---

### 4. Row & Column Operations

Manage heights, widths, visibility, grouping, auto-fit, and insert/delete operations.

```csharp
var sheet = workbook.ActiveWorksheet;

// Row properties
ExcelRow row = sheet.Row(1);
row.Height = 30.0;
row.Hidden = false;

// Column properties
ExcelColumn col = sheet.Column(1); // Column 'A'
col.Width = 25.0;
col.AutoFit();                     // Auto-fit column width based on content

// Insert & Delete Rows / Columns
sheet.InsertRows(startRow: 2, count: 5);
sheet.DeleteRows(startRow: 10, count: 2);

sheet.InsertColumns(startColumn: 2, count: 1);
sheet.DeleteColumns(startColumn: 5, count: 1);

// Group & Ungroup Rows / Columns
sheet.GroupRows(fromRow: 2, toRow: 10);
sheet.UngroupRows(fromRow: 2, toRow: 10);

sheet.GroupColumns(fromColumn: 1, toColumn: 3);
sheet.UngroupColumns(fromColumn: 1, toColumn: 3);

// First & Last Used Row / Column
ExcelRow? firstRow = sheet.FirstRowUsed();
ExcelRow? lastRow  = sheet.LastRowUsed();
ExcelColumn? firstCol = sheet.FirstColumnUsed();
ExcelColumn? lastCol  = sheet.LastColumnUsed();
```

---

### 5. Cell Reading & Writing (Type-Safe Getters)

Set and read cell values with type safety and fallback default values.

```csharp
// Cell assignment via address string or row/column numbers
sheet.Cell("A1").Value = "Product Name";
sheet.Cell(1, 2).Value = 199.99;
sheet["A3"].Value = DateTime.UtcNow;
sheet[1, 4].Value = true;

// Type-safe getters
ExcelCell cell = sheet.Cell("B1");

string strVal  = cell.GetString();
int intVal     = cell.GetInt32();
long longVal   = cell.GetInt64();
double dblVal  = cell.GetDouble();
decimal decVal = cell.GetDecimal();
bool boolVal   = cell.GetBoolean();
DateTime dtVal = cell.GetDateTime();

// Generic getter & TryGetValue with fallback
if (cell.HasValue)
{
    double val = cell.GetValue<double>();
    
    if (cell.TryGetValue<int>(out int parsedCount))
    {
        Console.WriteLine($"Count: {parsedCount}");
    }
}

int fallback = cell.GetValueOrDefault<int>(defaultValue: 0);
```

---

### 6. Range Operations & Bulk Matrix Access

Perform bulk formatting, matrix assignments, clearing, merging, and copying.

```csharp
var range = sheet.Range("A1:D10");
// Or range by coordinates: sheet.Range(startRow: 1, startColumn: 1, endRow: 10, endColumn: 4);

// Merge & Unmerge
range.Merge();
range.Unmerge();

// Clear Operations
range.ClearContents(); // Clears values and formulas
range.ClearStyles();   // Resets styles to default
range.Clear();         // Clears values, formulas, styles, comments, hyperlinks

// Bulk Copy
var targetRange = sheet.Range("F1:I10");
range.CopyTo(targetRange);

// Bulk Matrix Value Assignment & Extraction
object?[,] matrix = new object?[,]
{
    { "ID", "Name", "Price" },
    { 101, "Keyboard", 89.99 },
    { 102, "Mouse", 29.99 }
};
sheet.SetValues(matrix, startRow: 1, startColumn: 1);

object?[,] extracted = sheet.GetValues(startRow: 1, startColumn: 1, endRow: 3, endColumn: 3);
```

---

### 7. Cell Styling & Number Formatting

Customize fonts, fills, borders, alignments, and number format strings.

```csharp
using SmartExcelKit.Styles;

var customStyle = new ExcelStyle(
    font: new ExcelFont(name: "Arial", size: 12, bold: true, italic: false, underline: true, color: "1F4E79"),
    fill: new ExcelFill(ExcelFillPatternType.Solid, backgroundColor: "EBF1F5"),
    border: new ExcelBorder(
        left: new ExcelBorderItem(ExcelBorderStyle.Thin, color: "000000"),
        right: new ExcelBorderItem(ExcelBorderStyle.Thin, color: "000000"),
        top: new ExcelBorderItem(ExcelBorderStyle.Medium, color: "1F4E79"),
        bottom: new ExcelBorderItem(ExcelBorderStyle.Medium, color: "1F4E79")
    ),
    alignment: new ExcelAlignment(
        horizontal: ExcelHorizontalAlignment.Center,
        vertical: ExcelVerticalAlignment.Center,
        wrapText: true,
        indent: 1
    ),
    numberFormat: new ExcelNumberFormat("$#,##0.00")
);

// Apply style to cell or range
sheet.Cell("B2").Style = customStyle;
sheet.Range("A1:D1").Style = customStyle;
```

---

### 8. Excel Tables (ListObject) & Totals Row

Create native Excel Tables with custom styles and totals row aggregations.

```csharp
using SmartExcelKit.Tables;

var table = sheet.Tables.Add("A1:C10", "SalesTable");

// Table Properties
table.StyleName = "TableStyleMedium9";
table.ShowHeaderRow = true;
table.ShowTotalsRow = true;
table.ShowStripedRows = true;
table.ShowStripedColumns = false;

// Configure Totals Row Functions for Columns
table["Price"].TotalsRowFunction = TotalsRowFunction.Average;
table["Quantity"].TotalsRowFunction = TotalsRowFunction.Sum;
table["ID"].TotalsRowFunction = TotalsRowFunction.Count;
```

---

### 9. AutoFilter & Multi-Column Sorting

Enable AutoFilters and perform multi-column, culture-aware sorting.

```csharp
// Enable AutoFilter over specific range or used range
sheet.AutoFilter("A1:D100");
sheet.ClearAutoFilter();

// Multi-column sorting: sort rows 2-100 by Column 2 (ascending)
sheet.Sort(
    startRow: 2, startColumn: 1,
    endRow: 100, endColumn: 4,
    sortColumn: 2,
    ascending: true,
    culture: System.Globalization.CultureInfo.CurrentCulture
);
```

---

### 10. Conditional Formatting

Apply color scales, data bars, value comparisons, or formula rules.

```csharp
using SmartExcelKit.Formatting;
using SmartExcelKit.Styles;

// 1. Cell Value Rule (Highlight values > 500 in red)
var alertStyle = new ExcelStyle(font: new ExcelFont(bold: true, color: "FF0000"));
sheet.ConditionalFormatting.AddCellValueRule("C2:C100", ConditionalFormattingOperator.GreaterThan, "500", alertStyle);

// 2. Data Bar Gradient Visualization
sheet.ConditionalFormatting.AddDataBar("D2:D100", colorHex: "00FF00");

// 3. 2-Color & 3-Color Scales
sheet.ConditionalFormatting.AddTwoColorScale("E2:E100", minColorHex: "FFFFFF", maxColorHex: "0000FF");
sheet.ConditionalFormatting.AddThreeColorScale("F2:F100", minColorHex: "FF0000", midColorHex: "FFFF00", maxColorHex: "00FF00");

// 4. Formula Rule
sheet.ConditionalFormatting.AddFormulaRule("A2:A100", "ISODD(ROW())", alertStyle);
```

---

### 11. Data Validation Rules

Add drop-down lists, numeric ranges, text length limits, date constraints, or custom formulas.

```csharp
using SmartExcelKit.Validation;

// 1. List Validation (Drop-down menu)
sheet.DataValidations.AddListValidation("A2:A100", new[] { "Approved", "Pending", "Rejected" });

// 2. Whole Number & Decimal Validation
sheet.DataValidations.AddWholeNumberValidation("B2:B100", ValidationOperator.Between, min: 1, max: 100);
sheet.DataValidations.AddDecimalValidation("C2:C100", ValidationOperator.GreaterThan, min: 0.0);

// 3. Date & Text Length Validation
sheet.DataValidations.AddDateValidation("D2:D100", ValidationOperator.GreaterThanOrEqual, minDate: DateTime.Today);
sheet.DataValidations.AddTextLengthValidation("E2:E100", ValidationOperator.LessThanOrEqual, maxLen: 50);

// 4. Custom Formula Validation & Alert Configuration
var customVal = sheet.DataValidations.AddCustomValidation("F2:F100", "=AND(F2>0, F2<1000)");
customVal.ShowErrorMessage = true;
customVal.ErrorTitle = "Invalid Entry";
customVal.ErrorMessage = "Value must be between 0 and 1000.";
```

---

### 12. Formula Engine & Recalculation

Evaluate 60+ built-in functions with dependency tracking and calculation caching.

```csharp
using SmartExcelKit.Formula;

sheet["A1"].Value = 100;
sheet["A2"].Value = 200;
sheet["A3"].Formula = "SUM(A1:A2)";
sheet["A4"].Formula = "AVERAGE(A1:A3)";
sheet["A5"].Formula = "VLOOKUP(100, A1:A2, 1, FALSE)";

// Recalculate dependent cell formulas across the entire workbook
workbook.Recalculate();

// Evaluate formula directly without placing it in a cell
object? evalResult = FormulaEvaluator.Evaluate("MAX(A1:A2) * 1.1", sheet);
Console.WriteLine($"Result: {evalResult}");
```

---

### 13. Rich Text, Hyperlinks & Cell Comments

Attach formatted text runs, hyperlinks, and styled notes to cells.

```csharp
using SmartExcelKit.Core;

var cell = sheet["A1"];

// 1. Rich Text
var rich = new RichText();
rich.AddBold("Status: ", fontSize: 11, fontColorHex: "0000FF");
rich.AddItalic("Action Required", fontSize: 10, fontColorHex: "FF0000");
cell.RichText = rich;

// 2. Hyperlinks (Internal, External, Email, File)
cell.HyperlinkObject = ExcelHyperlink.External("https://github.com", tooltip: "Open Web Site");
cell.HyperlinkObject = ExcelHyperlink.Internal("Sheet2", "B5", tooltip: "Jump to Sheet2");
cell.HyperlinkObject = ExcelHyperlink.Email("support@example.com", subject: "Inquiry");

// 3. Cell Comments
cell.CommentObject = new ExcelComment("Please verify this figure before publishing.", author: "Financial Controller");
```

---

### 14. Images, Native Charts & Pivot Tables

Embed visual elements directly into worksheets.

```csharp
using SmartExcelKit.Drawings;
using SmartExcelKit.Core;

// 1. Embedded Image (PNG, JPEG, SVG, GIF, BMP)
var img = ExcelImage.FromFile("logo.png", topRow: 1, leftColumn: 5, width: 200, height: 100);
sheet.Images.Add(img);

// 2. Native Chart (Column, Bar, Line, Pie, Area, Scatter, Doughnut)
var chart = new ExcelChart(ChartType.Column, topRow: 5, leftColumn: 5, width: 450, height: 300);
chart.Title = "Quarterly Revenue";
chart.AddSeries("Sales", "Sheet1!B2:B10");
sheet.Charts.Add(chart);

// 3. Embedded Pivot Table
var pivot = new ExcelPivotTable("SalesPivot", "Sheet1!A1:D100", targetCell: new CellAddress(15, 5));
pivot.AddRowField("Category");
pivot.AddColumnField("Year");
pivot.AddDataField("Revenue", PivotSummaryFunction.Sum);
pivot.AddFilterField("Region");
sheet.PivotTables.Add(pivot);
```

---

### 15. Page Setup & Print Settings

Configure orientation, paper size, print areas, margins, and headers/footers.

```csharp
using SmartExcelKit.PageSetup;

var setup = sheet.PageSetup;
setup.Orientation = PageOrientation.Landscape;
setup.PaperSize = PaperSize.A4;
setup.PrintArea = "A1:G50";
setup.PrintGridlines = true;
setup.PrintHeadings = true;

// Fit to Pages
setup.FitToPages = true;
setup.FitToWidth = 1;
setup.FitToHeight = 2;

// Headers & Footers
setup.HeaderCenter = "Quarterly Financial Report";
setup.FooterRight = "Page 1 of 10";

// Margins (in inches)
setup.MarginLeft = 0.75;
setup.MarginRight = 0.75;
setup.MarginTop = 1.0;
setup.MarginBottom = 1.0;
```

---

### 16. Named Ranges (Workbook & Worksheet Scoped)

Create and manage named ranges globally or scoped to a specific worksheet.

```csharp
using SmartExcelKit.Core;

// Workbook-scoped named range
workbook.NamedRanges.Add("GlobalTaxRate", "Sheet1!$Z$1");

// Worksheet-scoped named range
sheet.NamedRanges.Add("LocalData", "Sheet1!$A$1:$D$50");

// Access named range
ExcelNamedRange taxRange = workbook.NamedRanges["GlobalTaxRate"];
```

---

### 17. High-Performance Streaming Reader & Writer

Process giant XLSX files with 1,000,000+ rows under 15 MB RAM.

```csharp
using SmartExcelKit.Streaming;

// 1. Streaming Writer (Flat memory write)
using var outStream = File.Create("large_export.xlsx");
using var writer = new ExcelStreamingWriter(outStream);

writer.BeginSheet("Data");
writer.WriteRow(new object?[] { "ID", "User", "Timestamp" });

// Sync & Async row streaming
for (int i = 1; i <= 1000000; i++)
{
    writer.WriteRow(new object?[] { i, $"User_{i}", DateTime.UtcNow });
}
await writer.WriteRowAsync(new object?[] { 1000001, "FinalUser", DateTime.UtcNow });

writer.EndSheet();

// 2. Streaming Reader (Sync & Async forward-only read)
using var inStream = File.OpenRead("large_export.xlsx");
using var reader = new ExcelStreamingReader(inStream);

foreach (string sheetName in reader.GetSheets())
{
    foreach (object?[] rowValues in reader.ReadRows(sheetName))
    {
        // Process row without memory overhead
    }

    List<object?[]> allRowsAsync = await reader.ReadRowsAsync(sheetName);
}
```

---

### 18. Delimited CSV & TSV Engine

High-speed character buffer streaming parser for CSV/TSV files.

```csharp
using SmartExcelKit.Csv;

// 1. Detect Encoding & Delimiter automatically
using var readStream = File.OpenRead("data.csv");
var encoding = CsvEngine.DetectEncoding(readStream);
char delimiter = CsvEngine.DetectDelimiter(readStream, encoding);

// 2. Read Streaming
foreach (List<string> rowFields in CsvEngine.ReadStreaming(readStream, delimiter, encoding))
{
    Console.WriteLine($"Col 0: {rowFields[0]}, Col 1: {rowFields[1]}");
}

// 3. Write CSV
var rows = new List<List<string>>
{
    new() { "ID", "Name" },
    new() { "1", "Alice" }
};
using var writeStream = File.Create("export.csv");
CsvEngine.Write(writeStream, rows, delimiter: ',');
```

---

### 19. POCO & DataTable Integration

Import C# object collections or DataTables into worksheets, or export them back.

```csharp
using System.Data;

public class Product
{
    public string Name { get; set; } = string.Empty;
    public double Price { get; set; }
    public int Quantity { get; set; }
}

var products = new List<Product>
{
    new() { Name = "Laptop", Price = 999.99, Quantity = 5 },
    new() { Name = "Mouse", Price = 25.50, Quantity = 40 }
};

// 1. POCO Import & Export
sheet.Import(products, startRow: 1, startColumn: 1, includeHeader: true);
List<Product> exportedProducts = sheet.Export<Product>(startRow: 1, startColumn: 1).ToList();

// 2. DataTable Import & Export
var dt = new DataTable("Orders");
dt.Columns.Add("ID", typeof(int));
dt.Columns.Add("Customer", typeof(string));
dt.Rows.Add(1, "Acme Corp");

sheet.Import(dt, startRow: 1, startColumn: 1, includeHeader: true);
DataTable exportedDt = sheet.ExportToDataTable(startRow: 1, startColumn: 1, hasHeader: true);
```

---

### 20. Developer One-Liner Export & Conversion Helpers

Quickly convert entire workbooks, worksheets, or rows directly to CSV, JSON, DataSet, DataTable, Dictionaries, or POCO collections without writing manual loops.

```csharp
using System.Data;

// 1. Workbook-Level One-Liners (Aggregates across ALL worksheets in the workbook)
string fullWorkbookCsv  = workbook.ToCsv(delimiter: ',');
string fullWorkbookJson = workbook.ToJson(hasHeader: true);
DataSet workbookDataSet = workbook.ToDataSet(hasHeader: true);    // DataSet containing DataTables per sheet
DataTable activeSheetDt = workbook.ToDataTable(hasHeader: true);
List<Product> allPocos  = workbook.ToObjects<Product>().ToList();   // Maps POCOs across all sheets
var allSheetDicts       = workbook.ToDictionaryList();             // List of dictionaries with '__Worksheet' key

// 2. Worksheet-Level One-Liners
string csvString   = sheet.ToCsv(delimiter: ',');
string jsonString  = sheet.ToJson(hasHeader: true);
DataTable dataTbl  = sheet.ToDataTable(hasHeader: true);
var dictList       = sheet.ToDictionaryList(startRow: 1);
List<Product> list = sheet.ToObjects<Product>().ToList();

// 3. Row-Level & Cell Helpers
foreach (var row in sheet.RowsUsed())
{
    string rowCsv                 = row.ToCsv();
    object?[] rowValues           = row.Values();
    Dictionary<string, object?> d = row.ToDictionary(headerRow: 1);
    Product p                     = row.ToObject<Product>(headerRow: 1);
    bool blank                    = row.IsBlank();
    ExcelCell? firstCell          = row.FirstCellUsed();
    ExcelCell? lastCell           = row.LastCellUsed();
}
```

---

## License

This project is licensed under the **MIT License**.