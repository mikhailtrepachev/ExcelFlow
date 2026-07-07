# ExcelFlow

[![NuGet Version](https://img.shields.io/nuget/v/ExcelFlow.svg)](https://www.nuget.org/packages/ExcelFlow/)
[![.NET Native AOT Ready](https://img.shields.io/badge/.NET-Native_AOT_Ready-blue.svg)](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**ExcelFlow** is a high-performance, strictly typed, and memory-efficient Excel (.xlsx) reader and writer designed specifically for modern .NET. 

Built from the ground up to support **Native AOT** and **Assembly Trimming**, ExcelFlow eliminates runtime reflection entirely by leveraging Roslyn Source Generators. It uses forward-only streaming to process hundreds of thousands of rows with a minimal memory footprint.

## 🚀 Key Features

* **100% Native AOT & Zero-Allocation:** Zero reflection and zero boxing at runtime. Parsers and Writers are fully generated at compile time as strictly typed instructions without any garbage allocations on the hot path.
* **Extreme Performance & Low Memory:** Built on top of `ExcelDataReader` and `OpenXml` with a strict forward-only streaming approach. Read and write massive datasets without loading the entire file into RAM.
* **Fluent Data Validation:** Define business rules inline. Filter out invalid rows automatically.
* **Graceful Error Handling:** Catch parsing or type-conversion errors via `.OnError()` callbacks instead of dealing with random runtime exceptions.
* **Skip Rows & Layout Resilience:** Easily skip title rows or descriptive headers before parsing the actual table structure.
* **Asynchronous Streaming:** First-class support for `IAsyncEnumerable<T>` to stream rows asynchronously (ideal for web APIs and microservices).

## 📦 Installation

```bash
dotnet add package ExcelFlow
```

## 📖 Quick Start

### 1. Define your model

Mark your class with `[ExcelFlowSerializable]` to trigger the Source Generator. Map properties to Excel columns using `[ExcelColumn]`.
You can map columns by Name, by Index, and define whether they are required.

```csharp
using ExcelFlow;

namespace MyApp;

[ExcelFlowSerializable]
public partial class SalesRecord
{
    [ExcelColumn("Transaction ID")]
    public int TransactionId { get; set; }

    [ExcelColumn("Product Name", IsRequired = true)] // Fail-fast if missing
    public string? ProductName { get; set; }

    [ExcelColumn(Index = 2)] // Direct mapping by column index (0-based)
    public decimal Amount { get; set; }

    [ExcelColumn("Date")]
    public DateTime Date { get; set; }
}
```
> **Note:** The class must be `partial` to allow the Source Generator to extend it.

### 2. Reading Data

Read the data synchronously into a `List<T>`, or stream it asynchronously with rich validation and error handling.

```csharp
using ExcelFlow;

var records = await Excel.Read<SalesRecord>("sales_data.xlsx")
    .FromSheet("Q1_Sales")
    .SkipRows(2) // Skip title and description rows before the header
    .Validate(row => row.Amount >= 0, "Amount cannot be negative")
    .OnError(error => 
    {
        Console.WriteLine($"[Warning] Row {error.RowNumber}: Column '{error.ColumnName}' failed. " +
                          $"Expected {error.ExpectedType}, but got '{error.RawValue}'.");
    })
    .ToListAsync(); // Or use .AsAsyncEnumerable() for true streaming
```

### 3. Writing Data

Export collections to Excel instantly with zero boxing and zero reflection.

```csharp
using ExcelFlow;

List<SalesRecord> data = GetSalesData();

Excel.Write(data)
     .ToSheet("Exported_Sales")
     .ToFile("report.xlsx");
```

## 🛠 Advanced: True Async Streaming

If you are building a memory-constrained worker service or web endpoint, you can process rows one-by-one as they are parsed, keeping your RAM usage completely flat regardless of file size:

```csharp
await using var fileStream = File.OpenRead("huge_dataset.xlsx");

var asyncStream = Excel.Read<SalesRecord>(fileStream)
    .AsAsyncEnumerable(cancellationToken);

await foreach (var record in asyncStream)
{
    await _database.InsertAsync(record);
}
```

## 💉 Dependency Injection (ASP.NET Core)

ExcelFlow provides a built-in interface `IExcelFlowService` for easy Dependency Injection, making your code perfectly testable.

```csharp
// Program.cs
builder.Services.AddExcelFlow();
```

Inject and use it in your controllers or services:

```csharp
public class ReportController
{
    private readonly IExcelFlowService _excel;

    public ReportController(IExcelFlowService excel)
    {
        _excel = excel;
    }

    public async Task<IActionResult> Upload(IFormFile file)
    {
        var data = await _excel.Read<SalesRecord>(file.OpenReadStream()).ToListAsync();
        return Ok(data);
    }
}
```

## 📄 License

ExcelFlow is licensed under the MIT License. See the [LICENSE](LICENSE.md) file for more details.
