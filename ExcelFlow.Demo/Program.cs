using System.Diagnostics;

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
        string filePath = "excelflow_aot_demo.xlsx";
        string sheetName = "SalesData";
        int rowCount = 500000;

        Console.WriteLine("==================================================");
        Console.WriteLine("🚀 EXCELFLOW 1.1.0 : NATIVE AOT DEMO");
        Console.WriteLine("==================================================\n");

        // ==========================================
        // 1. DATA GENERATION (EXPORT)
        // ==========================================
        Console.WriteLine($"[1/3] Generating Excel file with {rowCount:N0} rows...");

        IEnumerable<SalesRow> fakeData = Enumerable.Range(1, rowCount).Select(i => new SalesRow
        {
            Manager = $"Manager_{i}",
            Amount = i % 10 == 0 ? -50m : i * 100.5m, // Introduce some negative invalid amounts
            Date = DateTime.Now.AddDays(-i)
        });

        Stopwatch sw = Stopwatch.StartNew();

        // Fluent API Export (Zero Reflection)
        Excel.Write(fakeData)
             .ToSheet(sheetName)
             .ToFile(filePath);

        Console.WriteLine($" ✔ File generated in {sw.ElapsedMilliseconds} ms.");
        Console.WriteLine($" ✔ File Size: {new FileInfo(filePath).Length / 1024 / 1024} MB\n");

        // ==========================================
        // 2. STANDARD READ (IN-MEMORY LIST)
        // ==========================================
        Console.WriteLine("[2/3] Reading file into Memory (List<T>) with validation...");
        RunGarbageCollector();
        long memBeforeList = GC.GetTotalMemory(true);
        sw.Restart();

        int errorCount = 0;

        // Fluent API Read with Graceful Error Handling and AOT validation
        List<SalesRow> rowsFromFile = await Excel.Read<SalesRow>(filePath)
                                      .FromSheet(sheetName)
                                      .Validate(row => row.Amount > 0, "Amount must be strictly positive")
                                      .OnError(err => errorCount++)
                                      .ToListAsync();

        long timeList = sw.ElapsedMilliseconds;
        long memAfterList = GC.GetTotalMemory(false);
        PrintResult("ExcelFlow (ToList)", timeList, memAfterList - memBeforeList);
        Console.WriteLine($"   > Valid Rows: {rowsFromFile.Count:N0} | Rejected Rows: {errorCount:N0}");

        // ==========================================
        // 3. ASYNC STREAMING (LOW MEMORY FOOTPRINT)
        // ==========================================
        Console.WriteLine("\n[3/3] Streaming file asynchronously (IAsyncEnumerable)...");
        RunGarbageCollector();
        long memBeforeStream = GC.GetTotalMemory(true);
        sw.Restart();

        decimal totalAmount = 0;
        await using var fileStream = File.OpenRead(filePath);

        var asyncStream = Excel.Read<SalesRow>(fileStream)
                               .FromSheet(sheetName)
                               .Validate(row => row.Amount > 0, "Amount must be strictly positive")
                               .AsAsyncEnumerable();

        // Processing rows one by one without loading the whole file into RAM
        await foreach (var row in asyncStream)
        {
            totalAmount += row.Amount;
        }

        long timeStream = sw.ElapsedMilliseconds;
        long memAfterStream = GC.GetTotalMemory(false);
        PrintResult("ExcelFlow (Async Stream)", timeStream, memAfterStream - memBeforeStream);
        Console.WriteLine($"   > Processed Total Amount: {totalAmount:C2}");

        // ==========================================
        // SUMMARY
        // ==========================================
        Console.WriteLine("\n📊 AOT PERFORMANCE SUMMARY:");

        long listMemoryMb = Math.Max(1, (memAfterList - memBeforeList) / 1024 / 1024);
        long streamMemoryMb = Math.Max(0, (memAfterStream - memBeforeStream) / 1024 / 1024);

        Console.WriteLine($" - List Parsing Time: {timeList} ms");
        Console.WriteLine($" - Stream Parsing Time: {timeStream} ms");
        Console.WriteLine($" - RAM Used (List): {listMemoryMb} MB (Objects stored in memory)");
        Console.WriteLine($" - RAM Used (Stream): {streamMemoryMb} MB (True streaming power!)");
        Console.WriteLine($" - Native AOT Compatibility: 100% (No reflection used)");

        if (File.Exists(filePath)) File.Delete(filePath);

        Console.WriteLine("\nDemo complete.");
        if (Environment.GetEnvironmentVariable("CI") != "true")
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }

    static void PrintResult(string name, long timeMs, long memoryBytes)
    {
        double memoryMb = Math.Max(0, memoryBytes / (1024.0 * 1024.0));
        Console.WriteLine($"   > Method: {name}");
        Console.WriteLine($"   > Time: {timeMs} ms");
        Console.WriteLine($"   > Memory Allocated: {memoryMb:F2} MB");
    }

    static void RunGarbageCollector()
    {
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
    }
}