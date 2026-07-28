using System.Data;
using SmartExcelKit.Core;
using SmartExcelKit.Csv;
using SmartExcelKit.Drawings;
using SmartExcelKit.Formatting;
using SmartExcelKit.PageSetup;
using SmartExcelKit.Streaming;
using SmartExcelKit.Styles;
using SmartExcelKit.Tables;
using SmartExcelKit.Validation;

namespace SmartExcelKit.Sample;

/// <summary>
/// Sample application demonstrating SmartExcelKit features and generating an Excel workbook for testing.
/// </summary>
public static class Program
{
    /// <summary>
    /// Entry point of the sample application.
    /// </summary>
    public static async Task Main()
    {
        Console.WriteLine("==================================================================");
        Console.WriteLine("        SmartExcelKit Enterprise Library - Comprehensive Sample    ");
        Console.WriteLine("==================================================================");

        try
        {
            // 1. Generate Master Excel File containing all features for Excel inspection
            string masterFilePath = GenerateMasterExcelDemoFile();

            // 2. Additional streaming & CSV engine demonstrations
            await DemoStreamingReaderAndWriterAsync();
            DemoCsvEngine();

            Console.WriteLine("\n==================================================================");
            Console.WriteLine("    ALL DEMONSTRATIONS COMPLETED SUCCESSFULLY WITHOUT ERRORS!      ");
            Console.WriteLine("==================================================================");
            Console.WriteLine($"\n[SUCCESS] Open the master Excel file below in Microsoft Excel:");
            Console.WriteLine($"-> FILE PATH: {masterFilePath}\n");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n[ERROR] An unexpected error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }

    /// <summary>
    /// Generates a comprehensive Excel Workbook (.xlsx) with dedicated worksheets for every feature.
    /// </summary>
    public static string GenerateMasterExcelDemoFile()
    {
        Console.WriteLine("\nGenerating Master Excel File ('SmartExcelKit_Master_Demo.xlsx')...");
        using var workbook = new ExcelWorkbook();

        // -------------------------------------------------------------
        // Sheet 1: Data Types, AutoFit & Range Operations
        // -------------------------------------------------------------
        var s1 = workbook.AddWorksheet("01_DataTypes");
        s1.Cell("A1").Value = "Data Type"; s1.Cell("B1").Value = "Sample Value";
        s1.Range("A1:B1").Style = GetHeaderStyle();

        s1.Cell("A2").Value = "String Text"; s1.Cell("B2").Value = "SmartExcelKit High-Performance Library";
        s1.Cell("A3").Value = "Integer"; s1.Cell("B3").Value = 1250;
        s1.Cell("A4").Value = "Double"; s1.Cell("B4").Value = 99.95;
        s1.Cell("A5").Value = "Boolean"; s1.Cell("B5").Value = true;
        s1.Cell("A6").Value = "DateTime"; s1.Cell("B6").Value = new DateTime(2026, 7, 28);

        // AutoFit Column Width Demonstration
        s1.Range("A1:B6").AutoFitColumns();

        // -------------------------------------------------------------
        // Sheet 2: Cell Styling, Merged Cells (Satır/Hücre Birleştirme) & Column Width
        // -------------------------------------------------------------
        var s2 = workbook.AddWorksheet("02_Styles_Formatting");

        // Merged Row Header Demonstration (Satır/Hücre Birleştirme)
        s2.Cell("A1").Value = "QUARTERLY SALES PERFORMANCE SUMMARY";
        s2.MergeCells("A1:D1");
        s2.Row(1).Height = 35;
        s2.Cell("A1").Style = new ExcelStyle(
            font: new ExcelFont(name: "Calibri", size: 14, bold: true, color: "FFFFFF"),
            fill: new ExcelFill(ExcelFillPatternType.Solid, backgroundColor: "1F4E79"),
            alignment: new ExcelAlignment(horizontal: ExcelHorizontalAlignment.Center, vertical: ExcelVerticalAlignment.Center)
        );

        s2.Cell("A2").Value = "Product"; s2.Cell("B2").Value = "Category"; s2.Cell("C2").Value = "Price"; s2.Cell("D2").Value = "Margin";
        s2.Range("A2:D2").Style = GetHeaderStyle();

        s2.Cell("A3").Value = "Laptop Pro"; s2.Cell("B3").Value = "Hardware"; s2.Cell("C3").Value = 1299.99; s2.Cell("D3").Value = 0.25;
        s2.Cell("A4").Value = "Wireless Mouse"; s2.Cell("B4").Value = "Accessories"; s2.Cell("C4").Value = 29.50; s2.Cell("D4").Value = 0.40;

        // Custom Number Format & Alignment
        var currencyStyle = new ExcelStyle(numberFormat: new ExcelNumberFormat("$#,##0.00"), alignment: new ExcelAlignment(horizontal: ExcelHorizontalAlignment.Right));
        var percentStyle = new ExcelStyle(numberFormat: new ExcelNumberFormat("0.0%"), alignment: new ExcelAlignment(horizontal: ExcelHorizontalAlignment.Right));

        s2.Cell("C3").Style = currencyStyle; s2.Cell("C4").Style = currencyStyle;
        s2.Cell("D3").Style = percentStyle; s2.Cell("D4").Style = percentStyle;

        // Explicit Column Width & AutoFit
        s2.Column(1).AutoFit();
        s2.Column(2).Width = 20.0;
        s2.Column(3).Width = 18.0;
        s2.Column(4).Width = 15.0;

        // -------------------------------------------------------------
        // Sheet 3: Excel Tables, AutoFilter & Sorting
        // -------------------------------------------------------------
        var s3 = workbook.AddWorksheet("03_Tables_AutoFilter");
        s3.Cell("A1").Value = "Employee ID"; s3.Cell("B1").Value = "Department"; s3.Cell("C1").Value = "Salary";
        s3.Cell("A2").Value = 101; s3.Cell("B2").Value = "Engineering"; s3.Cell("C2").Value = 85000.0;
        s3.Cell("A3").Value = 102; s3.Cell("B3").Value = "Marketing"; s3.Cell("C3").Value = 62000.0;
        s3.Cell("A4").Value = 103; s3.Cell("B4").Value = "Engineering"; s3.Cell("C4").Value = 94000.0;

        var table = s3.Tables.Add("A1:C4", "EmployeeTable");
        table.StyleName = "TableStyleMedium2";
        table.ShowTotalsRow = true;
        table["Salary"].TotalsRowFunction = TotalsRowFunction.Sum;

        s3.AutoFilter("A1:C4");
        s3.Sort(startRow: 2, startColumn: 1, endRow: 4, endColumn: 3, sortColumn: 3, ascending: false);
        s3.Range("A1:C4").AutoFitColumns();

        // -------------------------------------------------------------
        // Sheet 4: Conditional Formatting Rules
        // -------------------------------------------------------------
        var s4 = workbook.AddWorksheet("04_ConditionalFormatting");
        s4.Cell("A1").Value = "Values (>500 Alert)"; s4.Cell("B1").Value = "Data Bar"; s4.Cell("C1").Value = "2-Color Scale"; s4.Cell("D1").Value = "3-Color Scale";
        s4.Range("A1:D1").Style = GetHeaderStyle();

        for (int r = 2; r <= 11; r++)
        {
            int val = (r - 1) * 100;
            s4.Cell(r, 1).Value = val;
            s4.Cell(r, 2).Value = val;
            s4.Cell(r, 3).Value = val;
            s4.Cell(r, 4).Value = val;
        }

        var alertStyle = new ExcelStyle(font: new ExcelFont(bold: true, color: "FF0000"));
        s4.ConditionalFormatting.AddCellValueRule("A2:A11", ConditionalFormattingOperator.GreaterThan, "500", alertStyle);
        s4.ConditionalFormatting.AddDataBar("B2:B11", "00FF00");
        s4.ConditionalFormatting.AddColorScale2("C2:C11", "FFFFFF", "0000FF");
        s4.ConditionalFormatting.AddThreeColorScale("D2:D11", "FF0000", "FFFF00", "00FF00");
        s4.Range("A1:D11").AutoFitColumns();

        // -------------------------------------------------------------
        // Sheet 5: Data Validation Rules
        // -------------------------------------------------------------
        var s5 = workbook.AddWorksheet("05_DataValidation");
        s5.Cell("A1").Value = "Dropdown List"; s5.Cell("B1").Value = "Integer (1-100)"; s5.Cell("C1").Value = "Date (>= Today)"; s5.Cell("D1").Value = "Text Len (<=10)"; s5.Cell("E1").Value = "Custom Formula";
        s5.Range("A1:E1").Style = GetHeaderStyle();

        s5.DataValidations.AddListValidation("A2:A10", new string[] { "Option A", "Option B", "Option C" });
        s5.DataValidations.AddWholeNumberValidation("B2:B10", ValidationOperator.Between, min: 1, max: 100);
        s5.DataValidations.AddDateValidation("C2:C10", ValidationOperator.GreaterThanOrEqual, minDate: DateTime.Today);
        s5.DataValidations.AddTextLengthValidation("D2:D10", ValidationOperator.LessThanOrEqual, minLength: 1, maxLength: 10);

        var customVal = s5.DataValidations.AddCustomValidation("E2:E10", "=AND(E2>0, E2<500)");
        customVal.ShowErrorMessage = true;
        customVal.ErrorTitle = "Invalid Range";
        customVal.ErrorMessage = "Entered value must be between 1 and 499.";
        s5.Range("A1:E10").AutoFitColumns();

        // -------------------------------------------------------------
        // Sheet 6: Advanced Formula Engine & Recalculate
        // -------------------------------------------------------------
        var s6 = workbook.AddWorksheet("06_Formulas");
        s6.Cell("A1").Value = "Item"; s6.Cell("B1").Value = "Qty"; s6.Cell("C1").Value = "UnitPrice"; s6.Cell("D1").Value = "SubTotal"; s6.Cell("E1").Value = "Status";
        s6.Range("A1:E1").Style = GetHeaderStyle();

        s6.Cell("A2").Value = "Keyboard"; s6.Cell("B2").Value = 5; s6.Cell("C2").Value = 49.99; s6.Cell("D2").Formula = "B2*C2"; s6.Cell("E2").Formula = "IF(D2>200, \"High Volume\", \"Regular\")";
        s6.Cell("A3").Value = "Monitor"; s6.Cell("B3").Value = 2; s6.Cell("C3").Value = 199.50; s6.Cell("D3").Formula = "B3*C3"; s6.Cell("E3").Formula = "IF(D3>200, \"High Volume\", \"Regular\")";
        s6.Cell("A4").Value = "Mouse Pad"; s6.Cell("B4").Value = 10; s6.Cell("C4").Value = 15.00; s6.Cell("D4").Formula = "B4*C4"; s6.Cell("E4").Formula = "IF(D4>200, \"High Volume\", \"Regular\")";

        s6.Cell("A6").Value = "Total Revenue"; s6.Cell("D6").Formula = "SUM(D2:D4)";
        s6.Cell("A7").Value = "Average Order"; s6.Cell("D7").Formula = "AVERAGE(D2:D4)";
        s6.Cell("A8").Value = "Rounded Average"; s6.Cell("D8").Formula = "ROUND(AVERAGE(D2:D4), 2)";
        s6.Cell("A9").Value = "Uppercase Item 1"; s6.Cell("D9").Formula = "UPPER(A2)";
        s6.Cell("A10").Value = "Absolute Variance"; s6.Cell("D10").Formula = "ABS(-142.50)";
        s6.Cell("A11").Value = "VLOOKUP Match Price"; s6.Cell("D11").Formula = "VLOOKUP(\"Keyboard\", A2:D4, 3, FALSE)";

        workbook.Recalculate();
        s6.Range("A1:E11").AutoFitColumns();

        // -------------------------------------------------------------
        // Sheet 7: Rich Text, Cell Comments & External/Internal Hyperlinks
        // -------------------------------------------------------------
        var s7 = workbook.AddWorksheet("07_RichText_Comments");
        s7.Cell("A1").Value = "Feature"; s7.Cell("B1").Value = "Sample Output";
        s7.Range("A1:B1").Style = GetHeaderStyle();

        // Rich Text Run
        var rich = new RichText();
        rich.AddBold("SmartExcelKit: ", fontSize: 13, fontColorHex: "1F4E79");
        rich.AddItalic("High Performance ", fontSize: 11, fontColorHex: "008000");
        rich.AddText(".NET Excel Library");
        s7.Cell("A2").Value = "Rich Text Multi-Run";
        s7.Cell("B2").RichText = rich;

        // External Web Link
        s7.Cell("A3").Value = "External Web Link";
        s7.Cell("B3").Value = "SmartExcelKit GitHub Web Repo";
        s7.Cell("B3").HyperlinkObject = ExcelHyperlink.External("https://github.com/erbasaran/SmartExcelKit.NET", tooltip: "Open Web Repo in Browser");

        // Internal Sheet Link
        s7.Cell("A4").Value = "Internal Sheet Link";
        s7.Cell("B4").Value = "Jump to Data Types Sheet";
        s7.Cell("B4").HyperlinkObject = ExcelHyperlink.Internal("01_DataTypes", "A1", tooltip: "Navigate to Sheet 1");

        // Email Link
        s7.Cell("A5").Value = "Email Link";
        s7.Cell("B5").Value = "Contact Support Email";
        s7.Cell("B5").HyperlinkObject = ExcelHyperlink.Email("support@smartexcelkit.net", subject: "SmartExcelKit Inquiry");

        // Cell Comment (Popup Hover Note)
        s7.Cell("A6").Value = "Cell Comment Note";
        s7.Cell("B6").Value = "Hover over this cell";
        s7.Cell("B6").CommentObject = new ExcelComment("This is a rich cell comment annotated with SmartExcelKit.", author: "Lead Architect");

        s7.Range("A1:B6").AutoFitColumns();

        // -------------------------------------------------------------
        // Sheet 8: Native Charts, Images & Pivot Tables
        // -------------------------------------------------------------
        var s8 = workbook.AddWorksheet("08_Charts_Images_Pivot");
        s8.Cell("A1").Value = "Category"; s8.Cell("B1").Value = "Region"; s8.Cell("C1").Value = "Year"; s8.Cell("D1").Value = "Revenue"; s8.Cell("E1").Value = "Units";
        s8.Cell("A2").Value = "Software"; s8.Cell("B2").Value = "North"; s8.Cell("C2").Value = "2025"; s8.Cell("D2").Value = 50000.0; s8.Cell("E2").Value = 120;
        s8.Cell("A3").Value = "Software"; s8.Cell("B3").Value = "South"; s8.Cell("C3").Value = "2025"; s8.Cell("D3").Value = 45000.0; s8.Cell("E3").Value = 110;
        s8.Cell("A4").Value = "Hardware"; s8.Cell("B4").Value = "North"; s8.Cell("C4").Value = "2026"; s8.Cell("D4").Value = 80000.0; s8.Cell("E4").Value = 200;
        s8.Cell("A5").Value = "Hardware"; s8.Cell("B5").Value = "South"; s8.Cell("C5").Value = "2026"; s8.Cell("D5").Value = 95000.0; s8.Cell("E5").Value = 240;

        // Native Column Chart
        var columnChart = new ExcelChart(ChartType.Column, topRow: 7, leftColumn: 1, width: 500, height: 300)
        {
            Title = "Regional Revenue Comparison",
            CategoryRange = "08_Charts_Images_Pivot!B2:B5",
            LegendPosition = LegendPosition.Right
        };
        columnChart.AddSeries("Revenue ($)", "08_Charts_Images_Pivot!D2:D5");
        s8.Charts.Add(columnChart);

        // Native Line Chart
        var lineChart = new ExcelChart(ChartType.Line, topRow: 7, leftColumn: 8, width: 500, height: 300)
        {
            Title = "Units Sold Trend",
            CategoryRange = "08_Charts_Images_Pivot!A2:A5",
            LegendPosition = LegendPosition.Bottom
        };
        lineChart.AddSeries("Units Sold", "08_Charts_Images_Pivot!E2:E5");
        s8.Charts.Add(lineChart);

        // Pivot Table
        var pivot = new ExcelPivotTable("SalesPivot", "08_Charts_Images_Pivot!A1:E5", targetCell: new CellAddress(23, 1));
        pivot.AddRowField("Category");
        pivot.AddColumnField("Year");
        pivot.AddDataField("Revenue", PivotSummaryFunction.Sum, customName: "Total Revenue ($)");
        pivot.AddDataField("Units", PivotSummaryFunction.Average, customName: "Avg Units");
        pivot.AddFilterField("Region");
        s8.PivotTables.Add(pivot);

        // Embedded Image (Valid 100x40 PNG Logo)
        byte[] samplePng = Convert.FromBase64String("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");
        s8.Images.Add(new ExcelImage(samplePng, ImageFormatType.Png, topRow: 1, leftColumn: 8, width: 220, height: 70) { Name = "Logo" });

        // -------------------------------------------------------------
        // Sheet 9: Page Setup & Printing
        // -------------------------------------------------------------
        var s9 = workbook.AddWorksheet("09_PageSetup");
        s9.PageSetup.Orientation = PageOrientation.Landscape;
        s9.PageSetup.PaperSize = PaperSize.A4;
        s9.PageSetup.PrintArea = "A1:F30";
        s9.PageSetup.PrintGridlines = true;
        s9.PageSetup.PrintHeadings = true;
        s9.PageSetup.HeaderCenter = "SmartExcelKit Official Executive Report";
        s9.PageSetup.FooterRight = "Page 1 of 1";
        s9.Cell("A1").Value = "Print Setup Configured (Landscape, A4, Print Gridlines & Headings Enabled).";

        // -------------------------------------------------------------
        // Sheet 10: POCO & DataTable Import
        // -------------------------------------------------------------
        var s10 = workbook.AddWorksheet("10_POCO_DataTable");
        var products = new List<SampleProductItem>
        {
            new() { Name = "Enterprise License", Stock = 50, Price = 499.00 },
            new() { Name = "Developer License", Stock = 200, Price = 199.00 }
        };
        s10.Import(products, startRow: 1, startColumn: 1);

        var dt = new DataTable("Orders");
        dt.Columns.Add("ID", typeof(int));
        dt.Columns.Add("Customer", typeof(string));
        dt.Rows.Add(1001, "Acme International");
        dt.Rows.Add(1002, "Global Tech Corp");
        s10.Import(dt, startRow: 6, startColumn: 1, includeHeader: true);
        s10.Range("A1:C10").AutoFitColumns();

        // -------------------------------------------------------------
        // Save Master Workbook to Output Folder and Project Root
        // -------------------------------------------------------------
        string outputDirectory = AppDomain.CurrentDomain.BaseDirectory;
        string masterFilePath = Path.Combine(outputDirectory, "SmartExcelKit_Master_Demo.xlsx");
        workbook.Save(masterFilePath);

        try
        {
            string rootPath = Path.GetFullPath(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "..", "SmartExcelKit_Master_Demo.xlsx"));
            workbook.Save(rootPath);
            Console.WriteLine($"-> Master Excel File generated at root: {rootPath}");
        }
        catch
        {
            // Ignore if root path fails
        }

        Console.WriteLine($"-> Master Excel File generated at: {masterFilePath}");
        return masterFilePath;
    }

    private static ExcelStyle GetHeaderStyle()
    {
        return new ExcelStyle(
            font: new ExcelFont(name: "Calibri", size: 11, bold: true, color: "FFFFFF"),
            fill: new ExcelFill(ExcelFillPatternType.Solid, backgroundColor: "1F4E79"),
            alignment: new ExcelAlignment(horizontal: ExcelHorizontalAlignment.Center, vertical: ExcelVerticalAlignment.Center)
        );
    }

    private static async Task DemoStreamingReaderAndWriterAsync()
    {
        Console.WriteLine("\nDemonstrating Low-Memory Streaming Reader & Writer...");

        string streamFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "stream_demo.xlsx");

        using (var outFs = File.Create(streamFile))
        using (var writer = new ExcelStreamingWriter(outFs))
        {
            writer.BeginSheet("StreamData");
            writer.WriteRow(new object?[] { "ID", "Name", "Timestamp" });

            for (int i = 1; i <= 50; i++)
            {
                writer.WriteRow(new object?[] { i, $"StreamingUser_{i}", DateTime.UtcNow });
            }
            await writer.WriteRowAsync(new object?[] { 51, "AsyncUser", DateTime.UtcNow });
            writer.EndSheet();
        }

        using (var inFs = File.OpenRead(streamFile))
        using (var reader = new ExcelStreamingReader(inFs))
        {
            foreach (string sheetName in reader.GetSheets())
            {
                var rows = await reader.ReadRowsAsync(sheetName);
                Console.WriteLine($"-> Streaming Reader read {rows.Count} rows from '{sheetName}'.");
            }
        }

        if (File.Exists(streamFile)) File.Delete(streamFile);
    }

    private static void DemoCsvEngine()
    {
        Console.WriteLine("\nDemonstrating CSV Engine & Encoding Auto-Detection...");

        string csvPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "csv_demo.csv");
        var rows = new List<List<string>>
        {
            new() { "ID", "ProductDescription" },
            new() { "1", "Ergonomic Mechanical Keyboard, RGB" },
            new() { "2", "High-Precision Laser Mouse" }
        };

        using (var fs = File.Create(csvPath))
        {
            CsvEngine.Write(fs, rows, delimiter: ',');
        }

        using (var fs = File.OpenRead(csvPath))
        {
            var encoding = CsvEngine.DetectEncoding(fs);
            char delimiter = CsvEngine.DetectDelimiter(fs, encoding);
            var readRows = CsvEngine.ReadStreaming(fs, delimiter, encoding).ToList();
            Console.WriteLine($"-> CsvEngine read {readRows.Count} rows (Encoding: {encoding.EncodingName}, Delimiter: '{delimiter}').");
        }

        if (File.Exists(csvPath)) File.Delete(csvPath);
    }
}

/// <summary>
/// Sample POCO product item representation.
/// </summary>
public sealed class SampleProductItem
{
    /// <summary>Gets or sets the product name.</summary>
    public string Name { get; set; } = string.Empty;

    /// <summary>Gets or sets the stock count.</summary>
    public int Stock { get; set; }

    /// <summary>Gets or sets the unit price.</summary>
    public double Price { get; set; }
}
