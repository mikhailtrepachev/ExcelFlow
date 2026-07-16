using ClosedXML.Excel;

namespace ExcelFlow.Benchmarks;

internal static class TestData
{
    public const int RowCount = 50_000;

    private static readonly string[] Categories =
        { "Electronics", "Books", "Toys", "Food", "Garden", "Auto", "Health", "Sports" };

    public static List<BenchRow> CreateRows(int count)
    {
        // Fixed seed: identical data in every benchmark process
        var rnd = new Random(42);
        var baseDate = new DateTime(2024, 1, 1);
        var list = new List<BenchRow>(count);

        for (int i = 0; i < count; i++)
        {
            list.Add(new BenchRow
            {
                Id = i + 1,
                Product = "Product " + rnd.Next(0, 500),
                Amount = Math.Round((decimal)(rnd.NextDouble() * 10_000), 2),
                Date = baseDate.AddDays(rnd.Next(0, 365)),
                Category = Categories[rnd.Next(Categories.Length)],
            });
        }

        return list;
    }

    /// <summary>
    /// Generates the shared read-benchmark file once. Written with ClosedXML so the file is
    /// "canonical" xlsx (shared strings, date styles) and no reader gets its own writer's quirks.
    /// </summary>
    public static string EnsureFile(int count)
    {
        string dir = Path.Combine(Path.GetTempPath(), "ExcelFlowBench");
        Directory.CreateDirectory(dir);
        string path = Path.Combine(dir, $"data_{count}.xlsx");

        if (File.Exists(path))
            return path;

        List<BenchRow> rows = CreateRows(count);

        using var wb = new XLWorkbook();
        IXLWorksheet ws = wb.AddWorksheet("Data");

        ws.Cell(1, 1).Value = "Id";
        ws.Cell(1, 2).Value = "Product";
        ws.Cell(1, 3).Value = "Amount";
        ws.Cell(1, 4).Value = "Date";
        ws.Cell(1, 5).Value = "Category";

        ws.Cell(2, 1).InsertData(rows.Select(r => new object[] { r.Id, r.Product!, r.Amount, r.Date, r.Category! }));

        wb.SaveAs(path);
        return path;
    }
}
