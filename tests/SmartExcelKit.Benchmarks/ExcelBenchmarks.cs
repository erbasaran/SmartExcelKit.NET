using BenchmarkDotNet.Attributes;
using SmartExcelKit.Csv;
using SmartExcelKit.Formula;
using System.Text;

namespace SmartExcelKit.Benchmarks;

[MemoryDiagnoser]
public class ExcelBenchmarks : System.IDisposable
{
    private byte[]? _csvData;
    private ExcelWorkbook? _workbook;
    private ExcelWorksheet? _worksheet;

    public void Dispose()
    {
        _workbook?.Dispose();
    }

    [GlobalSetup]
    public void Setup()
    {
        // Setup large CSV data (1000 rows)
        var sb = new StringBuilder();
        sb.AppendLine("ID,Name,Value,Active");
        for (int i = 0; i < 1000; i++)
        {
            sb.AppendLine($"{i},Product_{i},{10.5 * i},true");
        }
        _csvData = Encoding.UTF8.GetBytes(sb.ToString());

        // Setup Workbook for formula evaluation
        _workbook = new ExcelWorkbook();
        _worksheet = _workbook.AddWorksheet("Sheet1");
        for (int i = 1; i <= 100; i++)
        {
            _worksheet.Cell(i, 1).Value = (double)i;
        }
    }

    [Benchmark]
    public int Benchmark_CsvParsing()
    {
        using var ms = new MemoryStream(_csvData!);
        var rows = CsvEngine.Read(ms, ',');
        return rows.Count;
    }

    [Benchmark]
    public object? Benchmark_FormulaEvaluation()
    {
        return FormulaEvaluator.Evaluate("=SUM(A1:A100) + 50", _worksheet!, new Core.CellAddress(101, 1));
    }

    [Benchmark]
    public void Benchmark_WorkbookSaveXlsx()
    {
        using var outMs = new MemoryStream();
        _workbook!.Save(outMs, ExcelFileFormat.Xlsx);
    }
}
