**ExcelFlow** is an ultra-fast, memory-efficient, and strictly typed Excel (`.xlsx`) reader and writer for modern .NET.

It provides an elegant, LINQ-like API similar to Entity Framework, but without the massive memory footprint of traditional DOM-based libraries like ClosedXML or EPPlus.

## 💡 Why ExcelFlow?

Traditional Excel libraries load the entire document into memory to build a DOM (Document Object Model). Reading or exporting a 500,000-row file can easily consume **1.5 GB of RAM** and crash your microservices with an `OutOfMemoryException`.

**ExcelFlow solves this by using:**
1. **Forward-Only Streaming:** It reads and writes data row-by-row. Memory consumption stays completely flat (around ~50 MB) regardless of file size.
2. **Expression Trees:** No slow Reflection in loops. Mapping is compiled into raw machine code at runtime, allowing you to process over **160,000 rows per second**.

## 📦 Installation

Install via NuGet Package Manager:
```bash
dotnet add package ExcelFlow
```

## 📖 Quick Start
1. Define your model
   Use the [ExcelColumn] attribute to map Excel headers to your C# properties. Supports Nullable types and robust Date parsing (OADate / ISO).

```
using ExcelFlow;

public class SalesRow
{
    [ExcelColumn("Manager Name")]
    public string Manager { get; set; }
    
    public decimal Amount { get; set; }
    
    public DateTime? Date { get; set; } // Nullable supported!
}
```
**2. Read from Excel (Streaming)**
Thanks to yield return, data is streamed lazily. You can use standard LINQ methods without loading the entire file into memory.

```
// Open from file path or any Stream (e.g., HTTP Request)
using var db = new ExcelContext("huge_report.xlsx");

// LINQ in action - memory footprint stays at ~50MB!
var topSales = db.Worksheet<SalesRow>("Sales")
                 .Where(x => x.Amount > 1000)
                 .ToList();
```
**Safe Parsing & Error Handling:**
Excel data is often dirty. You can pass a callback to gracefully catch parsing errors without crashing your app:

```
var errors = new List<ExcelParseError>();
var data = db.Worksheet<SalesRow>("Sales", onError: err => errors.Add(err)).ToList();

// err contains RowNumber, ColumnName, ExpectedType, and the RawValue!
```

**3. Write to Excel (Export)**
Export millions of rows straight from your database to an Excel file or HTTP Response Stream. No intermediate memory buffers.

```
var myData = new List<SalesRow> { ... };

// 1. Export directly to a file
myData.ToExcel("export.xlsx", sheetName: "2024 Report");

// 2. Or stream directly to an HTTP Response in ASP.NET Core
[HttpGet("download")]
public IActionResult DownloadReport()
{
    var query = _dbContext.Sales.AsNoTracking();
    var stream = new MemoryStream();
    
    query.ToExcel(stream, "Export");
    
    stream.Position = 0;
    return File(stream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "report.xlsx");
}
```

## 📊 Benchmarks
**Test case: 500,000 rows, 3 columns (String, Decimal, DateTime).**

⏳ Running ClosedXML...
Sum: 1312507875000.0
Time: 7371 ms
Memory (RAM): 491.77 MB

🚀 Running ExcelFlow...
Sum: 1312507875000.0
Time: 1362 ms
Memory (RAM): 1.76 MB

🏆 FINAL COMPARISON:
Speedup: 5.4x faster!
Memory saved: 278.7x less RAM!