# ExcelFlow

[![NuGet Version](https://img.shields.io/nuget/v/ExcelFlow.svg)](https://www.nuget.org/packages/ExcelFlow/)
[![.NET Native AOT Ready](https://img.shields.io/badge/.NET-Native_AOT_Ready-blue.svg)](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

**ExcelFlow** is a high-performance, strictly typed, and memory-efficient Excel (.xlsx) reader and writer designed specifically for modern .NET.

Built from the ground up to support **Native AOT** and **Assembly Trimming**, ExcelFlow eliminates runtime reflection entirely by leveraging Roslyn Source Generators. It uses forward-only streaming to process hundreds of thousands of rows with a minimal memory footprint.

## 🚀 Key Features

* **Native AOT & Zero Reflection:** Mappers are fully generated at compile time as strictly typed code. No reflection, no runtime code generation, no surprises after `dotnet publish /p:PublishAot=true`.
* **Compile-Time Diagnostics:** Mistakes surface as build errors/warnings (`EXF001`–`EXF003`), not as silent runtime failures.
* **High Performance & Low Memory:** Strict forward-only streaming for both reading and writing. Rows are processed one at a time — the dataset is never fully materialized unless you ask for a `List<T>`.
* **Fluent Data Validation:** Define business rules inline. Invalid rows are filtered out automatically.
* **Graceful Error Handling:** Catch parsing and type-conversion errors via `.OnError()` callbacks with **exact 1-based sheet row numbers**, instead of dealing with random runtime exceptions.
* **Skip Rows & Layout Resilience:** Easily skip title rows or descriptive headers before parsing the actual table structure.
* **Proper Date Cells:** Exported `DateTime` values are written with a real date format — Excel shows dates, not raw numbers.
* **Asynchronous Streaming:** First-class support for `IAsyncEnumerable<T>` (ideal for web APIs and microservices).

## 📦 Installation

```bash
dotnet add package ExcelFlow
```

The package ships the source generator as a Roslyn analyzer — no extra packages needed.

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

    [ExcelColumn("Status")]
    public OrderStatus Status { get; set; } // enums are supported

    [ExcelColumn("External ID")]
    public Guid? ExternalId { get; set; }   // so are Guid and nullable types
}
```
> **Note:** The class must be `partial` to allow the Source Generator to extend it. If you forget, the build fails with a clear `EXF001` error.

**Natively supported property types:** `string`, `bool`, all numeric primitives, `DateTime`, `Guid`, enums — plus their nullable variants. Anything else produces an `EXF002` build warning and falls back to `Convert.ChangeType` at runtime.

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
        // RowNumber is the real 1-based sheet row — open the file and jump straight to it
        Console.WriteLine($"[Warning] Row {error.RowNumber}: Column '{error.ColumnName}' failed. " +
                          $"Expected {error.ExpectedType}, but got '{error.RawValue}'.");
    })
    .ToListAsync(); // Or use .AsAsyncEnumerable() for true streaming
```

`.OnError()` can be called multiple times — all handlers are invoked.

### 3. Writing Data

Export collections to Excel with compile-time generated writers. `DateTime` cells get a proper date format automatically (date-only values render as dates, timestamps as date + time).

```csharp
using ExcelFlow;

List<SalesRecord> data = GetSalesData();

Excel.Write(data)
     .ToSheet("Exported_Sales")
     .ToFile("report.xlsx");
```

## 🛠 Advanced: Async Streaming

If you are building a memory-constrained worker service or web endpoint, you can process rows one-by-one as they are parsed, without materializing the whole dataset:

```csharp
await using var fileStream = File.OpenRead("huge_dataset.xlsx");

var asyncStream = Excel.Read<SalesRecord>(fileStream)
    .AsAsyncEnumerable(cancellationToken);

await foreach (var record in asyncStream)
{
    await _database.InsertAsync(record);
}
```

> **Memory note:** rows are streamed, but the xlsx shared-string table is loaded once per file (a property of the format itself). For files with millions of *unique* strings, expect the baseline memory to be proportional to the shared-string table, not to the row count.

### Stream ownership

Since v1.1.3 ExcelFlow **does not dispose your stream** — you own it (this matches .NET conventions, e.g. `IFormFile.OpenReadStream()` in ASP.NET Core). If you want ExcelFlow to dispose the stream when reading completes, opt in explicitly:

```csharp
var records = Excel.Read<SalesRecord>(stream, leaveOpen: false).ToList();
```

## 🧭 Compile-Time Diagnostics

| Id | Severity | Meaning |
|----|----------|---------|
| `EXF001` | Error | A class marked `[ExcelFlowSerializable]` (or its enclosing type) is not `partial` |
| `EXF002` | Warning | Property type is not natively supported; runtime falls back to `Convert.ChangeType` |
| `EXF003` | Error | Generic classes are not supported |

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

## 🆕 What's new in 1.1.3

* **Packaging fix:** the source generator now ships inside the NuGet package. If mappers were not generated for you on 1.1.x — this release fixes it.
* **Correct row numbers:** `ExcelParseError.RowNumber` now reports the real 1-based sheet row, including when `SkipRows()` is used.
* **Robust headers:** empty header cells no longer throw `NullReferenceException`.
* **Culture-stable dates:** string dates parse with `InvariantCulture` on any server locale.
* **Dates look like dates:** exported `DateTime` cells carry a date number format.
* **Stream ownership:** your streams are no longer disposed by the library (opt back in with `leaveOpen: false`).
* **New types:** `Guid` and enum properties supported natively; nullable numeric/date properties fixed in the writer.
* **Compile-time diagnostics:** `EXF001`–`EXF003`.
* **Shared files:** files opened by path use `FileShare.Read` — reading works even while the file is open in Excel.

## 🗂 Repository Layout

| Project | Purpose |
|---------|---------|
| `ExcelFlow` | The runtime library (net8.0 / net9.0 / net10.0). Packs the generator into `analyzers/dotnet/cs`. |
| `ExcelFlow.SourceGenerators` | Roslyn incremental source generator (netstandard2.0) that emits AOT-safe mappers. |
| `ExcelFlow.Tests` | xUnit test suite (uses ClosedXML as an independent reference implementation). |
| `ExcelFlow.Demo` | Native AOT demo app (`PublishAot=true`): generates, reads and streams 500k rows. |

## 🔨 Building & Testing

```bash
dotnet build
dotnet test ExcelFlow.Tests

# Pack the NuGet package (the generator is bundled automatically)
dotnet pack ExcelFlow/ExcelFlow.csproj -c Release

# Run the AOT demo
dotnet publish ExcelFlow.Demo -c Release
```

## 📄 License

ExcelFlow is licensed under the MIT License. See the [LICENSE](LICENSE.md) file for more details.
