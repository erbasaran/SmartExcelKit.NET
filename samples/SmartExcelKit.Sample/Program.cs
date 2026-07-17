using SmartExcelKit.Formula;
using SmartExcelKit.Styles;

namespace SmartExcelKit.Sample;

public static class Program
{
    public static void Main()
    {
        Console.WriteLine("==================================================");
        Console.WriteLine("SmartExcelKit Enterprise Library - Demonstration");
        Console.WriteLine("==================================================");

        try
        {
            // 1. Create a Workbook and Worksheet
            using var workbook = new ExcelWorkbook();
            var sheet = workbook.AddWorksheet("SalesData");

            // 2. Add cell data with styling
            Console.WriteLine("Adding cell data and styling...");
            sheet.Cell("A1").Value = "Product";
            sheet.Cell("B1").Value = "Quantity";
            sheet.Cell("C1").Value = "Price";
            sheet.Cell("D1").Value = "Total";

            // Apply styling to headers
            var headerStyle = new ExcelStyle(font: new ExcelFont(name: "Arial", size: 12, bold: true, color: "FF0000FF"));
            sheet.Range("A1:D1").Style = headerStyle;

            // Row 2 Data
            sheet.Cell("A2").Value = "Smart Keyboard";
            sheet.Cell("B2").Value = 5.0;
            sheet.Cell("C2").Value = 99.99;
            sheet.Cell("D2").Formula = "B2*C2";

            // Row 3 Data
            sheet.Cell("A3").Value = "Wireless Mouse";
            sheet.Cell("B3").Value = 12.0;
            sheet.Cell("C3").Value = 19.50;
            sheet.Cell("D3").Formula = "B3*C3";

            // Row 4 Data (Total summary)
            sheet.Cell("A4").Value = "Grand Total";
            sheet.Cell("D4").Formula = "SUM(D2:D3)";
            sheet.Cell("A4").Style = new ExcelStyle(font: new ExcelFont(bold: true));
            sheet.Cell("D4").Style = new ExcelStyle(font: new ExcelFont(bold: true));

            // Auto-fit columns
            sheet.Range("A1:D4").AutoFitColumns();

            // 3. Evaluate formulas inside the program
            Console.WriteLine("Evaluating formulas via internal Formula Engine...");
            var totalRow2Val = FormulaEvaluator.Evaluate(sheet.Cell("D2").Formula!, sheet, sheet.Cell("D2").Address);
            var grandTotalVal = FormulaEvaluator.Evaluate(sheet.Cell("D4").Formula!, sheet, sheet.Cell("D4").Address);

            Console.WriteLine($"-> Row 2 Total (B2*C2): {totalRow2Val}");
            Console.WriteLine($"-> Grand Total (SUM(D2:D3)): {grandTotalVal}");

            // 4. Save to multiple formats
            string basePath = AppDomain.CurrentDomain.BaseDirectory;
            string xlsxPath = Path.Combine(basePath, "demo.xlsx");
            string csvPath = Path.Combine(basePath, "demo.csv");
            string jsonPath = Path.Combine(basePath, "demo.json");
            string htmlPath = Path.Combine(basePath, "demo.html");
            string xmlPath = Path.Combine(basePath, "demo.xml");

            Console.WriteLine($"Saving files in Base Directory: {basePath}");

            workbook.Save(xlsxPath);
            Console.WriteLine("Saved: XLSX");

            workbook.Save(csvPath);
            Console.WriteLine("Saved: CSV");

            workbook.Save(jsonPath);
            Console.WriteLine("Saved: JSON");

            workbook.Save(htmlPath);
            Console.WriteLine("Saved: HTML");

            workbook.Save(xmlPath);
            Console.WriteLine("Saved: XML Spreadsheet 2003");

            // 5. Demonstrate POCO Import / Export
            Console.WriteLine("Demonstrating POCO collections Import/Export...");
            var importSheet = workbook.AddWorksheet("POCO_Demo");
            var products = new List<ProductItem>
            {
                new() { Name = "Laptop Pro", Stock = 15, Price = 1299.99 },
                new() { Name = "USB Hub", Stock = 120, Price = 24.50 },
                new() { Name = "HDMI Cable", Stock = 80, Price = 9.99 }
            };

            importSheet.Import(products);
            Console.WriteLine("POCO list imported to worksheet.");

            var exportedProducts = importSheet.Export<ProductItem>();
            Console.WriteLine("Exporting back from worksheet to POCO list:");
            foreach (var item in exportedProducts)
            {
                Console.WriteLine($"  - Name: {item.Name}, Stock: {item.Stock}, Price: ${item.Price}");
            }

            // 6. Read back XLSX file
            Console.WriteLine("Reading back and verifying XLSX file...");
            using var reOpenedWb = ExcelWorkbook.Open(xlsxPath);
            var reOpenedSheet = reOpenedWb.Worksheets[0];
            Console.WriteLine($"Workbook opened successfully! Worksheet count: {reOpenedWb.Worksheets.Count}");
            Console.WriteLine($"First sheet cell A2: {reOpenedSheet.Cell("A2").Value}");
            Console.WriteLine($"First sheet cell D4 Formula: {reOpenedSheet.Cell("D4").Formula}");

            Console.WriteLine("==================================================");
            Console.WriteLine("Demonstration completed successfully!");
            Console.WriteLine("==================================================");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }
    }
}

public sealed class ProductItem
{
    public string Name { get; set; } = string.Empty;
    public int Stock { get; set; }
    public double Price { get; set; }
}
