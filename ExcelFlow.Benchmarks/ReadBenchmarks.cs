using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using MiniExcelLibs;

namespace ExcelFlow.Benchmarks;

/// <summary>
/// Reads the same 50k-row xlsx into List&lt;BenchRow&gt; with every library.
/// All readers do the same job: typed rows, all 5 columns materialized.
/// </summary>
[MemoryDiagnoser]
[ShortRunJob]
public class ReadBenchmarks
{
    private string _file = null!;

    [GlobalSetup]
    public void Setup()
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        _file = TestData.EnsureFile(TestData.RowCount);
    }

    [Benchmark(Baseline = true)]
    public List<BenchRow> ExcelFlow() => Excel.Read<BenchRow>(_file).ToList();

    /// <summary>Raw ExcelDataReader loop — the backend ExcelFlow wraps; shows our own overhead.</summary>
    [Benchmark]
    public List<BenchRow> ExcelDataReader_Raw()
    {
        using FileStream stream = File.Open(_file, FileMode.Open, FileAccess.Read, FileShare.Read);
        using var reader = global::ExcelDataReader.ExcelReaderFactory.CreateReader(stream);

        reader.Read(); // header row

        var list = new List<BenchRow>();
        while (reader.Read())
        {
            list.Add(new BenchRow
            {
                Id = (int)reader.GetDouble(0),
                Product = reader.GetString(1),
                Amount = (decimal)reader.GetDouble(2),
                Date = reader.GetDateTime(3),
                Category = reader.GetString(4),
            });
        }

        return list;
    }

    [Benchmark]
    public List<BenchRow> Sylvan()
    {
        using var reader = global::Sylvan.Data.Excel.ExcelDataReader.Create(_file);

        var list = new List<BenchRow>();
        while (reader.Read())
        {
            list.Add(new BenchRow
            {
                Id = reader.GetInt32(0),
                Product = reader.GetString(1),
                Amount = reader.GetDecimal(2),
                Date = reader.GetDateTime(3),
                Category = reader.GetString(4),
            });
        }

        return list;
    }

    [Benchmark]
    public List<BenchRow> MiniExcel() => MiniExcelLibs.MiniExcel.Query<BenchRow>(_file).ToList();

    [Benchmark]
    public List<BenchRow> ClosedXML()
    {
        using var wb = new XLWorkbook(_file);
        IXLWorksheet ws = wb.Worksheet(1);

        var list = new List<BenchRow>();
        foreach (IXLRow row in ws.RowsUsed().Skip(1))
        {
            list.Add(new BenchRow
            {
                Id = row.Cell(1).GetValue<int>(),
                Product = row.Cell(2).GetString(),
                Amount = row.Cell(3).GetValue<decimal>(),
                Date = row.Cell(4).GetDateTime(),
                Category = row.Cell(5).GetString(),
            });
        }

        return list;
    }
}
