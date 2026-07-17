using BenchmarkDotNet.Running;

namespace SmartExcelKit.Benchmarks;

public static class Program
{
    public static void Main(string[] args)
    {
        BenchmarkRunner.Run<ExcelBenchmarks>();
    }
}
