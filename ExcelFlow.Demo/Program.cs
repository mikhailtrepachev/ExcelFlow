using System.Diagnostics;
using ClosedXML.Excel;
using ExcelFlow;

namespace ExcelFlow.Demo;

[ExcelFlowSerializable]
public partial class SalesRow
{
    [ExcelColumn("Manager Name")]
    public string? Manager { get; set; }
    
    [ExcelColumn("Amount")]
    public decimal Amount { get; set; }
    
    [ExcelColumn("Date")]
    public DateTime Date { get; set; }
}

class Program
{
    static async Task Main()
    {
        string filePath = "demo_benchmark.xlsx";
        string sheetName = "SalesData";
        int rowCount = 500000;
        
        Console.WriteLine("=== EXCELFLOW: DEMONSTRATION AND BENCHMARK ===\n");

        // ==========================================
        // 1. DATA GENERATION
        // ==========================================
        Console.WriteLine($"[1/4] 🏭 Generating file with {rowCount:N0} rows...");
        IEnumerable<SalesRow> fakeData = Enumerable.Range(1, rowCount).Select(i => new SalesRow
        {
            Manager = $"Manager_{i}",
            Amount = i * 100.5m,
            Date = DateTime.Now.AddDays(-i)
        });

        Stopwatch sw = Stopwatch.StartNew();

        // Using your Fluent API for writing
        Excel.Write(fakeData)
             .ToSheet(sheetName)
             .ToFile(filePath);

        Console.WriteLine($"✅ File generated in {sw.ElapsedMilliseconds} ms. Size: {new FileInfo(filePath).Length / 1024 / 1024} MB\n");

        // ==========================================
        // 2. READING VIA CLOSEDXML (Competitor)
        // ==========================================
        Console.WriteLine("[2/4] ⏳ Reading via ClosedXML (Warning: consumes high RAM)...");
        RunGarbageCollector();
        long memBeforeClosed = GC.GetTotalMemory(true);
        sw.Restart();

        decimal totalClosedXML = 0;
        using (XLWorkbook workbook = new XLWorkbook(filePath))
        {
            IXLWorksheet? sheet = workbook.Worksheet(sheetName);
            foreach (IXLRow? row in sheet.RowsUsed().Skip(1))
            {
                totalClosedXML += row.Cell(2).GetValue<decimal>();
            }
        }

        long timeClosedXML = sw.ElapsedMilliseconds;
        long memAfterClosed = GC.GetTotalMemory(false);
        PrintResult("ClosedXML", timeClosedXML, memAfterClosed - memBeforeClosed);

        // ==========================================
        // 3. READING VIA EXCELFLOW (FILE)
        // ==========================================
        Console.WriteLine("\n[3/4] 🚀 Reading file via ExcelFlow (Fluent API)...");
        RunGarbageCollector();
        long memBeforeFlowFile = GC.GetTotalMemory(true);
        sw.Restart();

        // Demonstrating elegant Fluent API with error handling
        List<SalesRow> rowsFromFile = await Excel.Read<SalesRow>(filePath)
                                      .FromSheet(sheetName)
                                      .OnError(err => Console.WriteLine($"Error in row {err.RowNumber}: {err.RawValue}"))
                                      .ToListAsync();

        decimal totalFlowFile = rowsFromFile.Sum(x => x.Amount);

        long timeFlowFile = sw.ElapsedMilliseconds;
        long memAfterFlowFile = GC.GetTotalMemory(false);
        PrintResult("ExcelFlow (File)", timeFlowFile, memAfterFlowFile - memBeforeFlowFile);

        // ==========================================
        // 4. READING VIA EXCELFLOW (ASYNCHRONOUS STREAM)
        // ==========================================
        Console.WriteLine("\n[4/4] 🌊 Streaming asynchronous reading via ExcelFlow...");
        RunGarbageCollector();
        long memBeforeFlowStream = GC.GetTotalMemory(true);
        sw.Restart();

        decimal totalFlowStream = 0;
        await using var fileStream = File.OpenRead(filePath);

        // Demonstrating the power of IAsyncEnumerable!
        var asyncStream = Excel.Read<SalesRow>(fileStream)
                               .FromSheet(sheetName)
                               .AsAsyncEnumerable();

        await foreach (var row in asyncStream)
        {
            totalFlowStream += row.Amount;
        }

        long timeFlowStream = sw.ElapsedMilliseconds;
        long memAfterFlowStream = GC.GetTotalMemory(false);
        PrintResult("ExcelFlow (Async Stream)", timeFlowStream, memAfterFlowStream - memBeforeFlowStream);

        // ==========================================
        // SUMMARY
        // ==========================================
        Console.WriteLine("\n🏆 FINAL COMPARISON (ClosedXML vs ExcelFlow Stream):");

        double speedup = timeFlowStream == 0 ? 0 : (double)timeClosedXML / timeFlowStream;
        long closedMemoryMb = (memAfterClosed - memBeforeClosed) / 1024 / 1024;
        long flowMemoryMb = Math.Max(1, (memAfterFlowStream - memBeforeFlowStream) / 1024 / 1024); // Avoid div by zero
        double memorySavings = closedMemoryMb == 0 ? 0 : (double)closedMemoryMb / flowMemoryMb;

        Console.WriteLine($"⚡ Speedup: {speedup:F1}x faster");
        Console.WriteLine($"💾 Memory savings: {memorySavings:F1}x less RAM ({closedMemoryMb} MB vs {flowMemoryMb} MB)!");

        // Clean up
        if (File.Exists(filePath)) File.Delete(filePath);
    }

    static void PrintResult(string name, long timeMs, long memoryBytes)
    {
        double memoryMb = Math.Max(0, memoryBytes / (1024.0 * 1024.0));
        Console.WriteLine($"   > Library: {name}");
        Console.WriteLine($"   > Time: {timeMs} ms");
        Console.WriteLine($"   > Memory (RAM): {memoryMb:F2} MB");
    }

    static void RunGarbageCollector()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}