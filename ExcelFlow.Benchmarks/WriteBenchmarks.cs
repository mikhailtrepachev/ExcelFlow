using BenchmarkDotNet.Attributes;
using ClosedXML.Excel;
using MiniExcelLibs;

namespace ExcelFlow.Benchmarks;

/// <summary>
/// Writes the same 50k in-memory rows to an xlsx file with every library.
/// </summary>
[MemoryDiagnoser]
[ShortRunJob]
public class WriteBenchmarks
{
    private List<BenchRow> _data = null!;
    private string _dir = null!;

    [GlobalSetup]
    public void Setup()
    {
        _data = TestData.CreateRows(TestData.RowCount);
        _dir = Path.Combine(Path.GetTempPath(), "ExcelFlowBench", "out");
        Directory.CreateDirectory(_dir);
    }

    [Benchmark(Baseline = true)]
    public void ExcelFlow()
    {
        Excel.Write(_data).ToSheet("Data").ToFile(Path.Combine(_dir, "excelflow.xlsx"));
    }

    [Benchmark]
    public void MiniExcel()
    {
        MiniExcelLibs.MiniExcel.SaveAs(Path.Combine(_dir, "miniexcel.xlsx"), _data, overwriteFile: true);
    }

    [Benchmark]
    public void ClosedXML()
    {
        using var wb = new XLWorkbook();
        IXLWorksheet ws = wb.AddWorksheet("Data");

        ws.Cell(1, 1).Value = "Id";
        ws.Cell(1, 2).Value = "Product";
        ws.Cell(1, 3).Value = "Amount";
        ws.Cell(1, 4).Value = "Date";
        ws.Cell(1, 5).Value = "Category";

        ws.Cell(2, 1).InsertData(_data.Select(r => new object[] { r.Id, r.Product!, r.Amount, r.Date, r.Category! }));

        wb.SaveAs(Path.Combine(_dir, "closedxml.xlsx"));
    }
}
