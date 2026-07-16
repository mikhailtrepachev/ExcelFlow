using ExcelFlow;

namespace ExcelFlow.Benchmarks;

/// <summary>
/// Benchmark model. Property names are used as column headers,
/// so the same class maps 1:1 in every library under test.
/// </summary>
[ExcelFlowSerializable]
public partial class BenchRow
{
    public int Id { get; set; }

    public string? Product { get; set; }

    public decimal Amount { get; set; }

    public DateTime Date { get; set; }

    public string? Category { get; set; }
}
