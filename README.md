# SmartExcelKit

[![NuGet Version](https://img.shields.io/nuget/v/SmartExcelKit.svg)](https://www.nuget.org/packages/SmartExcelKit)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

SmartExcelKit is a modern, high-performance, low-allocation .NET Standard 2.0 library designed for reading, writing, and evaluating spreadsheet files (XLSX, XLS, CSV, TSV, HTML, XML Spreadsheet 2003, JSON).

Built for enterprise-grade performance and low-memory profiles, it features a custom formula engine with dependency cycle verification, automatic format signatures detection, and memory-optimized streaming readers/writers to handle millions of cells efficiently.

---

## Target Frameworks

- **SmartExcelKit**: `.NET Standard 2.0` (Supports .NET Core 2.0+, .NET 5+, and .NET Framework 4.6.1+).

---

## Installation

```shell
dotnet add package SmartExcelKit
```

---

## Core Features & Code Samples

### 1. Document Creation & Basic Cell Formatting

Create workbooks, add worksheets, adjust cell properties, and apply styles (which are centrally deduplicated at save time to minimize file size).

```csharp
using SmartExcelKit;
using SmartExcelKit.Core;
using SmartExcelKit.Styles;

using var workbook = new ExcelWorkbook();
var sheet = workbook.AddWorksheet("Inventory");

// Set Cell Values & Formulas
sheet.Cell("A1").Value = "Product Name";
sheet.Cell("B1").Value = "Unit Price";
sheet.Cell("C1").Value = "Quantity";
sheet.Cell("D1").Value = "Total Value";

sheet.Cell("A2").Value = "Smart Keyboard";
sheet.Cell("B2").Value = 99.99;
sheet.Cell("C2").Value = 10;
sheet.Cell("D2").Formula = "B2*C2"; // Auto-linked inside dependency graph

// Style Cells and Ranges (Font, Fill, Border, Alignment, NumberFormat)
var headerStyle = new ExcelStyle(
    font: new ExcelFont(name: "Segoe UI", size: 11, bold: true, color: "FFFFFF"),
    fill: new ExcelFill(ExcelFillPatternType.Solid, backgroundColor: "1F4E79"),
    alignment: new ExcelAlignment(horizontal: ExcelHorizontalAlignment.Center, vertical: ExcelVerticalAlignment.Center)
);

// Applies style to the entire range
sheet.Range("A1:D1").Style = headerStyle;

// Format price as currency
var priceStyle = new ExcelStyle(numberFormat: new ExcelNumberFormat("$#,##0.00"));
sheet.Range("B2:B2").Style = priceStyle;

// Save workbook (format auto-detected via file extension)
workbook.Save("inventory.xlsx");
```

---

### 2. Worksheets Layout (Width, Height, Merging, AutoFit)

Adjust worksheet row/column rendering properties and merge layouts.

```csharp
// Define column widths and row heights (1-based indices)
sheet.SetColumnWidth(1, 25.5);
sheet.SetRowHeight(1, 30.0);

// Hide rows or columns
sheet.SetColumnHidden(3, true);
sheet.SetRowHidden(10, true);

// Auto-fit column width based on cell contents length
sheet.AutoFitColumn(1);

// Merge cell areas
sheet.MergeCells(ExcelRangeAddress.Parse("A5:B6"));
sheet.UnmergeCells(ExcelRangeAddress.Parse("A5:B6"));
```

---

### 3. Worksheet Protection & Security

Protect worksheets with a password using the Excel 16-bit XOR hashing algorithm.

```csharp
// Protect the sheet
sheet.Protect("StrongPassword123");
Console.WriteLine(sheet.IsProtected); // True
Console.WriteLine(sheet.ProtectionPasswordHash); // Outputs standard Excel XOR password hash in HEX (e.g., "DAA7")

// Unprotect the sheet
sheet.Unprotect();
```

---

### 4. VBA Macro Preservation (.xlsm)

Opening a macro-enabled workbook preserves the `vbaProject.bin` structure and content-type associations on export.

```csharp
// Load macro-enabled spreadsheet (Macros/VBA preserved automatically)
using var workbook = ExcelWorkbook.Open("input_with_macros.xlsm");

// Make edits
workbook.Worksheets[0].Cell("A1").Value = "Updated Data";

// Save back as macro-enabled
workbook.Save("output_with_macros.xlsm");
```

---

### 5. Auto-Format & Flexible Loading (Path, Stream, Byte Array)

SmartExcelKit automatically detects file format content signatures (XLSX, XLS, CSV, HTML, JSON, XML Spreadsheet 2003) and supports loading from file paths, streams, or raw byte arrays.

```csharp
// 1. Load from file path (format auto-detected via extension/signature)
using var workbook1 = ExcelWorkbook.Open("data.xlsx");

// 2. Load from a stream (format auto-detected via signature)
using var stream = File.OpenRead("data.csv");
using var workbook2 = ExcelWorkbook.Open(stream);

// 3. Load from a raw byte array (format auto-detected via signature)
byte[] rawBytes = File.ReadAllBytes("data.xls"); // Works with legacy binary XLS
using var workbook3 = ExcelWorkbook.Open(rawBytes);
```

#### Legacy Excel Binary Formats (XLS & XLSB)
- **XLS (Excel 97-2003)**: Fully supported for **reading** natively out-of-the-box using our lightweight binary record parser (BIFF8/OLE2).
- **XLSB (Excel Binary)**: Reading is not directly supported in the core package. Attempting to load it throws a `WorkbookException` with code `LEGACY_BINARY_FORMAT_UNSUPPORTED`.

You can register custom XLSB providers via the format provider registry at application startup:

```csharp
using SmartExcelKit.Providers;

// Register custom XLSB format provider if needed
FormatProviderRegistry.Register(ExcelFileFormat.Xlsb, new MyCustomXlsbProvider());
```

---

### 6. High-Performance XLSX Streaming

For files containing millions of rows, use the forward-only streaming reader and writer to keep memory usage flat.

#### ExcelStreamingWriter (Flat Memory Write)
Writes rows directly to the ZIP stream using `inlineStr` nodes, avoiding memory allocations for string cache indexing.

```csharp
using var fs = File.Create("huge_export.xlsx");
using var writer = new ExcelStreamingWriter(fs);

writer.BeginSheet("HugeSheet");

// Write header row
writer.WriteRow(new object?[] { "ID", "Timestamp", "Data" });

// Stream data rows
for (int i = 1; i <= 1000000; i++)
{
    writer.WriteRow(new object?[] { i, DateTime.UtcNow, $"Data Item {i}" });
}

writer.EndSheet();
```

#### ExcelStreamingReader (Flat Memory Read)
Uses `XmlReader` underneath to yield worksheet rows on-the-fly.

```csharp
using var fs = File.OpenRead("huge_export.xlsx");
using var reader = new ExcelStreamingReader(fs);

foreach (string sheetName in reader.GetSheets())
{
    foreach (object?[] rowValues in reader.ReadRows(sheetName))
    {
        // rowValues indexes match columns (empty cells are aligned as nulls)
        Console.WriteLine($"Row: {string.Join(", ", rowValues)}");
    }
}
```

---

### 7. Zero-Allocation CSV Engine

Read and write raw delimited datasets utilizing buffer pools (`ArrayPool<char>`).

```csharp
using SmartExcelKit.Csv;

// 1. Write structured fields to CSV
var csvRows = new List<List<string>>
{
    new() { "Name", "Age", "Role" },
    new() { "Alice", "30", "Architect" },
    new() { "Bob, Developer", "25", "Developer" } // Auto-escaped with quotes
};
using var writeStream = File.Create("users.csv");
CsvEngine.Write(writeStream, csvRows, delimiter: ',');

// 2. Read fields streaming (low GC pressure)
using var readStream = File.OpenRead("users.csv");
var encoding = CsvEngine.DetectEncoding(readStream);
char delimiter = CsvEngine.DetectDelimiter(readStream, encoding);

foreach (List<string> row in CsvEngine.ReadStreaming(readStream, delimiter, encoding))
{
    Console.WriteLine($"Field 1: {row[0]}, Field 2: {row[1]}");
}
```

---

### 8. POCO & DataTable Data Binding

Import collections or `DataTable` instances directly to worksheets, or export worksheet records back to typed classes.

```csharp
public sealed class Employee
{
    public string FullName { get; set; } = string.Empty;
    public double Salary { get; set; }
}

var employees = new List<Employee>
{
    new() { FullName = "Sarah Connor", Salary = 95000 },
    new() { FullName = "John Connor", Salary = 60000 }
};

// Import to sheet
var sheet = workbook.AddWorksheet("Employees");
sheet.Import(employees);

// Export back from sheet to POCO collection
List<Employee> loaded = sheet.Export<Employee>().ToList();
```

---

### 9. Custom Excel Formula Engine

Parse and evaluate cell expressions locally without requiring Microsoft Excel. Supports range flattening and verification:
- **Math & Stats**: `SUM`, `AVERAGE`, `COUNT`, `COUNTA`, `MIN`, `MAX`
- **Logical**: `IF`, `AND`, `OR`, `NOT`
- **Text**: `LEFT`, `RIGHT`, `LEN`, `CONCAT`
- **Date & Time**: `TODAY`, `NOW`

```csharp
using SmartExcelKit.Formula;

// Setting formulas checks dependencies and guards against circular loops:
sheet.Cell("A1").Formula = "B1+1";
// sheet.Cell("B1").Formula = "A1+1"; // Throws FormulaException (Circular reference detected)

// Evaluate formula explicitly
object? result = FormulaEvaluator.Evaluate("SUM(A1:B10) * 10", sheet, CellAddress.Parse("C1"));
Console.WriteLine($"Result: {result}");
```

---

## License

This project is licensed under the **MIT License**.
