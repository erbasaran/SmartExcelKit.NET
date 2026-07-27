# SmartExcelKit

[![NuGet Version](https://img.shields.io/nuget/v/SmartExcelKit.svg)](https://www.nuget.org/packages/SmartExcelKit)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

SmartExcelKit is a modern, high-performance, low-allocation .NET Standard 2.0 library designed for reading, writing, and evaluating spreadsheet files (XLSX, XLS, CSV, TSV, HTML, XML Spreadsheet 2003, JSON).

---

## Target Frameworks

- **SmartExcelKit**: `.NET Standard 2.0` (Supports .NET Core 2.0+, .NET 5+, .NET 6/7/8/9/10+, and .NET Framework 4.6.1+).

---

## Installation

```shell
dotnet add package SmartExcelKit
```

---

## Complete API Guide & Examples

### 1. Open Workbook

Load workbooks seamlessly from file paths, streams, or raw byte arrays with automatic format detection.

```csharp
using SmartExcelKit;

// 1. Open from file path (auto-detects extension & signature)
using var workbook1 = ExcelWorkbook.Open("data.xlsx");

// 2. Open asynchronously from file path
using var workbook2 = await ExcelWorkbook.OpenAsync("data.xlsx");

// 3. Open from a stream
using var stream = File.OpenRead("data.csv");
using var workbook3 = ExcelWorkbook.Open(stream);

// 4. Open from byte array
byte[] fileBytes = File.ReadAllBytes("legacy.xls");
using var workbook4 = ExcelWorkbook.Open(fileBytes);
```

---

### 2. Create Workbook

Create new workbooks, add worksheets, set cell values and formulas, apply styles, and save to any supported format.

```csharp
using SmartExcelKit;
using SmartExcelKit.Styles;

using var workbook = new ExcelWorkbook();
var sheet = workbook.AddWorksheet("Inventory");

// Set cell values
sheet.Cell("A1").Value = "Product Name";
sheet.Cell("B1").Value = "Unit Price";
sheet.Cell("C1").Value = "Quantity";
sheet.Cell("D1").Value = "Total Value";

sheet.Cell("A2").Value = "Mechanical Keyboard";
sheet.Cell("B2").Value = 129.99;
sheet.Cell("C2").Value = 15;
sheet.Cell("D2").Formula = "B2*C2";

// Create & apply styled range
var headerStyle = new ExcelStyle(
    font: new ExcelFont(name: "Segoe UI", size: 11, bold: true, color: "FFFFFF"),
    fill: new ExcelFill(ExcelFillPatternType.Solid, backgroundColor: "1F4E79"),
    alignment: new ExcelAlignment(horizontal: ExcelHorizontalAlignment.Center)
);
sheet.Range("A1:D1").Style = headerStyle;

// Save workbook (format auto-detected via extension)
workbook.Save("inventory.xlsx");
```

---

### 3. Worksheet Navigation

Access sheets by name, 0-based index, or active worksheet reference.

```csharp
using var workbook = ExcelWorkbook.Open("data.xlsx");

// Access worksheet by name via indexer
ExcelWorksheet sheetByName = workbook["Inventory"];

// Access worksheet by 0-based index
ExcelWorksheet sheetByIndex = workbook[0];

// Access or set ActiveWorksheet
ExcelWorksheet activeSheet = workbook.ActiveWorksheet;
workbook.ActiveWorksheet = sheetByName;
```

---

### 4. Rows Enumeration

Lazily iterate all populated rows or access individual rows by 1-based index.

```csharp
var sheet = workbook.ActiveWorksheet;

// Lazy enumeration over all populated rows
foreach (ExcelRow row in sheet.Rows)
{
    Console.WriteLine($"Row {row.Index} - Height: {row.Height}, Hidden: {row.Hidden}, IsEmpty: {row.IsEmpty}");
}

// Access specific row by index
ExcelRow thirdRow = sheet.Row(3);
thirdRow.Height = 28.0;
thirdRow.Hidden = false;
```

---

### 5. Columns Enumeration

Lazily iterate all columns, query column letters, or modify widths and visibility.

```csharp
var sheet = workbook.ActiveWorksheet;

// Lazy enumeration over all populated columns
foreach (ExcelColumn col in sheet.Columns)
{
    Console.WriteLine($"Column {col.Letter} ({col.Index}) - Width: {col.Width}, IsEmpty: {col.IsEmpty}");
}

// Access specific column by index
ExcelColumn colB = sheet.Column(2); // Column 'B'
colB.Width = 22.5;
colB.AutoFit();
```

---

### 6. Cells Enumeration

Iterate cells sequentially over a range or worksheet without intermediate array allocations.

```csharp
var range = sheet.Range("A1:C10");

// Enumerates all cells in range row-by-row
foreach (ExcelCell cell in range.Cells)
{
    Console.WriteLine($"{cell.Address} = {cell.Value}");
}
```

---

### 7. UsedRange

Retrieve the bounding box of all populated cells in O(1) cached time.

```csharp
ExcelRange usedRange = sheet.UsedRange;

Console.WriteLine($"Used Range Address : {usedRange.Address.Address}");
Console.WriteLine($"Used Rows Count    : {sheet.UsedRowCount}");
Console.WriteLine($"Used Columns Count : {sheet.UsedColumnCount}");
Console.WriteLine($"Max Row            : {sheet.MaxRow}");
Console.WriteLine($"Max Column         : {sheet.MaxColumn}");
```

---

### 8. CellsUsed

Lazily enumerate only cells that contain actual values, formulas, comments, hyperlinks, or non-default styles.

```csharp
// Zero-allocation iteration over used cells only
foreach (ExcelCell cell in sheet.CellsUsed())
{
    Console.WriteLine($"Cell {cell.Address} ({cell.RowNumber}, {cell.ColumnNumber}) = {cell.Value}");
}
```

---

### 9. RowsUsed

Find the first, last, or all rows containing populated data.

```csharp
ExcelRow? firstRow = sheet.FirstRowUsed();
ExcelRow? lastRow  = sheet.LastRowUsed();

Console.WriteLine($"First row with data: {firstRow?.Index}");
Console.WriteLine($"Last row with data : {lastRow?.Index}");

// Iterate populated rows only
foreach (ExcelRow row in sheet.RowsUsed())
{
    Console.WriteLine($"Row {row.Index} contains {row.CellsUsed().Count()} used cells.");
}
```

---

### 10. ColumnsUsed

Find the first, last, or all columns containing populated data.

```csharp
ExcelColumn? firstCol = sheet.FirstColumnUsed();
ExcelColumn? lastCol  = sheet.LastColumnUsed();

Console.WriteLine($"First col with data: {firstCol?.Letter}");
Console.WriteLine($"Last col with data : {lastCol?.Letter}");

// Iterate populated columns only
foreach (ExcelColumn col in sheet.ColumnsUsed())
{
    Console.WriteLine($"Column {col.Letter} contains data.");
}
```

---

### 11. Range Operations

Perform bulk styling, value assignment, clear, merge, and copy operations across ranges.

```csharp
var range = sheet.Range("A1:D10");

// Bulk value assignment
range.Value = "Default Text";

// Merge & Unmerge
range.Merge();
range.Unmerge();

// Clear contents or styles
range.ClearContents(); // Clears values and formulas only
range.ClearStyles();   // Resets formatting to default
range.Clear();          // Clears everything (values, formulas, styles, comments, hyperlinks)

// Copy range to destination
var targetRange = sheet.Range("F1:I10");
range.CopyTo(targetRange);
```

---

### 12. Cell Access

Intuitive cell access via functions or indexers on worksheets, ranges, rows, and columns.

```csharp
// Method call access
sheet.Cell("A1").Value = "Title";
sheet.Cell(1, 1).Value = "Title";

// Indexer access on worksheet
sheet["A1"].Value = "Title";
sheet[1, 1].Value = "Title";

// Access via Row or Column wrappers
sheet.Row(1)[1].Value = "Title";
sheet.Column(1)[1].Value = "Title";
```

---

### 13. Typed Value Reading

Type-safe, null-safe helper methods and generic converters.

```csharp
ExcelCell cell = sheet.Cell("B2");

// Strongly typed convenience getters
int intVal       = cell.GetInt32();
long longVal     = cell.GetInt64();
double doubleVal = cell.GetDouble();
decimal decVal   = cell.GetDecimal();
bool boolVal     = cell.GetBoolean();
DateTime dateVal = cell.GetDateTime();
string strVal    = cell.GetString();

// Generic type conversion
if (cell.HasValue)
{
    double price = cell.GetValue<double>();
    
    if (cell.TryGetValue<int>(out int count))
    {
        Console.WriteLine($"Parsed count: {count}");
    }
}
```

---

### 14. Import POCO

Import strongly typed objects into worksheet rows with automatic header generation.

```csharp
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

var sheet = workbook.AddWorksheet("Products");
sheet.Import(products, startRow: 1, startColumn: 1);
```

---

### 15. Export POCO

Export sheet rows back to a strongly typed collection of POCOs matching header names.

```csharp
// Export worksheet rows back to POCO objects
List<Product> items = sheet.Export<Product>(startRow: 1, startColumn: 1).ToList();

foreach (var item in items)
{
    Console.WriteLine($"{item.Name}: ${item.Price} (Qty: {item.Quantity})");
}
```

---

### 16. Import DataTable

Import `DataTable` structures directly or export sheet content back to `DataTable`.

```csharp
using System.Data;

var dt = new DataTable("Orders");
dt.Columns.Add("OrderID", typeof(int));
dt.Columns.Add("Customer", typeof(string));
dt.Rows.Add(101, "Acme Corp");

// Import DataTable to sheet
sheet.Import(dt, startRow: 1, startColumn: 1, includeHeader: true);

// Export sheet content back to DataTable
DataTable exportedDt = sheet.ExportToDataTable(startRow: 1, startColumn: 1, hasHeader: true);
```

---

### 17. Formula Engine

Parse and evaluate Excel formulas with dependency cycle protection.

```csharp
using SmartExcelKit.Formula;

sheet.Cell("A1").Value = 50;
sheet.Cell("A2").Value = 150;
sheet.Cell("A3").Formula = "SUM(A1:A2)"; // Built-in formula linking

// Direct formula evaluation
object? result = FormulaEvaluator.Evaluate("AVERAGE(A1:A2) * 2", sheet);
Console.WriteLine($"Evaluated Result: {result}"); // 200
```

---

### 18. Worksheet Protection

Protect worksheets with password hashing compatible with standard Excel XOR protection.

```csharp
// Protect sheet with password
sheet.Protect("SecretPassword123");
Console.WriteLine(sheet.IsProtected); // True
Console.WriteLine(sheet.ProtectionPasswordHash); // HEX Hash e.g. "DAA7"

// Unprotect sheet
sheet.Unprotect();
```

---

### 19. Streaming Reader

Low-memory forward-only XML reader for processing massive files with millions of rows.

```csharp
using SmartExcelKit.Streaming;

using var fs = File.OpenRead("large_data.xlsx");
using var reader = new ExcelStreamingReader(fs);

foreach (string sheetName in reader.GetSheets())
{
    foreach (object?[] rowValues in reader.ReadRows(sheetName))
    {
        Console.WriteLine($"Row: {string.Join(", ", rowValues)}");
    }
}
```

---

### 20. Streaming Writer

Flat-memory forward-only ZIP stream writer that avoids string table allocations.

```csharp
using SmartExcelKit.Streaming;

using var fs = File.Create("large_export.xlsx");
using var writer = new ExcelStreamingWriter(fs);

writer.BeginSheet("Data");
writer.WriteRow(new object?[] { "ID", "Name", "Date" });

for (int i = 1; i <= 1000000; i++)
{
    writer.WriteRow(new object?[] { i, $"User {i}", DateTime.UtcNow });
}

writer.EndSheet();
```

---

### 21. CSV Engine

High-performance delimited parser utilizing pooled character memory (`ArrayPool<char>`).

```csharp
using SmartExcelKit.Csv;

// Write CSV
var rows = new List<List<string>>
{
    new() { "ID", "Description" },
    new() { "1", "Keyboard, Wireless" }
};
using var writeStream = File.Create("output.csv");
CsvEngine.Write(writeStream, rows, delimiter: ',');

// Read CSV with auto-detected encoding & delimiter
using var readStream = File.OpenRead("output.csv");
var encoding = CsvEngine.DetectEncoding(readStream);
char delimiter = CsvEngine.DetectDelimiter(readStream, encoding);

foreach (List<string> rowFields in CsvEngine.ReadStreaming(readStream, delimiter, encoding))
{
    Console.WriteLine($"Field 0: {rowFields[0]}, Field 1: {rowFields[1]}");
}
```

---

### 22. Auto Detect File Format

Automatically detect format from file extension or content magic signatures.

```csharp
using var stream = File.OpenRead("unknown_file");
ExcelFileFormat format = ExcelWorkbook.DetectFormat(stream, "unknown_file.xlsx");
Console.WriteLine($"Detected format: {format}"); // XLSX, XLS, CSV, JSON, XML2003, etc.
```

---

### 23. Bulk Operations

Matrix value manipulation, row/column insertion, deletion, and cell shifting.

```csharp
// 1. Bulk matrix assignment
var matrix = new object?[,]
{
    { "ID", "Name", "Score" },
    { 1, "Alice", 95.5 },
    { 2, "Bob", 88.0 }
};
sheet.SetValues(matrix, startRow: 1, startColumn: 1);

// 2. Retrieve matrix from region
object?[,] dataMatrix = sheet.GetValues(startRow: 1, startColumn: 1, endRow: 3, endColumn: 3);

// 3. Row & Column Shifting
sheet.InsertRows(startRow: 2, count: 1);   // Inserts 1 blank row at row 2
sheet.DeleteRows(startRow: 2, count: 1);   // Deletes 1 row at row 2

sheet.InsertColumns(startColumn: 2, count: 1); // Inserts 1 blank column at column B
sheet.DeleteColumns(startColumn: 2, count: 1); // Deletes 1 column at column B
```

---

### 25. Developer Helper & Convenience APIs

Write cleaner, highly expressive code without writing repetitive LINQ expressions or conversion boilerplate.

```csharp
// 1. One-liner CSV & JSON Exports
string sheetCsv = sheet.ToCsv(delimiter: ',');
string sheetJson = sheet.ToJson(hasHeader: true);

foreach (var row in sheet.RowsUsed())
{
    Console.WriteLine(row.ToCsv()); // Format row as CSV line
}

// 2. Direct Value Conversions & Dictionaries
object?[] values = row.Values();
Dictionary<string, object?> dict = row.ToDictionary(headerRow: 1);
List<Dictionary<string, object?>> dictList = sheet.ToDictionaryList();

// 3. One-liner Object Mapping
Person person = row.ToObject<Person>(headerRow: 1);
List<Person> people = sheet.ToObjects<Person>().ToList();
List<Person> rangePeople = sheet.Range("A1:D100").ToObjects<Person>().ToList();

// 4. Cell Helpers & Default Value Fallbacks
int age = cell.GetValueOrDefault<int>(defaultValue: 18);
string text = cell.AsString();
int count = cell.AsInt32();

// 5. Row Blank & Boundary Checking
bool isBlank = row.IsBlank();
ExcelCell? firstCell = row.FirstCellUsed();
ExcelCell? lastCell = row.LastCellUsed();
```

---

### 26. Performance Best Practices

To achieve optimal throughput and minimal GC pressure when working with large Excel files:

1. **Use `CellsUsed()`, `RowsUsed()`, `ColumnsUsed()`**: Avoid looping empty cells in large worksheets; `CellsUsed()` yields only populated coordinates.
2. **Prefer Lazy Enumeration**: Methods like `sheet.Rows` and `sheet.Columns` utilize `yield return` and do not allocate intermediate `List<T>` arrays.
3. **Use `ExcelStreamingWriter` & `ExcelStreamingReader`**: For datasets exceeding 100,000 rows, streaming reduces memory from hundreds of megabytes to a flat ~10 MB buffer.
4. **Cached Range Bounds**: `sheet.UsedRange`, `sheet.MaxRow`, and `sheet.MaxColumn` use internal O(1) bounds tracking, avoiding full dictionary scans on repeat calls.

---

## License

This project is licensed under the **MIT License**.