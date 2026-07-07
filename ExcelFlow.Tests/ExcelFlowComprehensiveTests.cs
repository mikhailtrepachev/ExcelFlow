using ClosedXML.Excel;

namespace ExcelFlow.Tests;

public partial class ExcelFlowComprehensiveTests
{
    // A fully compliant test model simulating an AOT-ready user class
    [ExcelFlowSerializable]
    public partial class ProductRecord
    {
        [ExcelColumn("ID")]
        public int Id { get; set; }
        
        [ExcelColumn("Product Name")]
        public string? Name { get; set; }
        
        [ExcelColumn("Price")]
        public decimal Price { get; set; }
        
        [ExcelColumn("Added")]
        public DateTime AddedDate { get; set; }
    }

    private Stream GenerateTestExcelStream(bool includeTitleRow = false)
    {
        var ms = new MemoryStream();
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Data");
            int startRow = 1;

            if (includeTitleRow)
            {
                ws.Cell(startRow, 1).Value = "Company Product Report 2026";
                startRow++;
            }

            // Headers
            ws.Cell(startRow, 1).Value = "ID";
            ws.Cell(startRow, 2).Value = "Product Name";
            ws.Cell(startRow, 3).Value = "Price";
            ws.Cell(startRow, 4).Value = "Added";

            // Data Row 1
            ws.Cell(startRow + 1, 1).Value = 1;
            ws.Cell(startRow + 1, 2).Value = "Apple";
            ws.Cell(startRow + 1, 3).Value = 1.50m;
            ws.Cell(startRow + 1, 4).Value = new DateTime(2026, 1, 1);

            // Data Row 2 (Invalid Price)
            ws.Cell(startRow + 2, 1).Value = 2;
            ws.Cell(startRow + 2, 2).Value = "Banana";
            ws.Cell(startRow + 2, 3).Value = "Free"; // Error expected here
            ws.Cell(startRow + 2, 4).Value = new DateTime(2026, 2, 1);

            wb.SaveAs(ms);
        }
        ms.Position = 0;
        return ms;
    }

    [Fact]
    public void Read_List_ParsesValidRows()
    {
        // Arrange
        using var stream = GenerateTestExcelStream();
        
        // Act
        var results = Excel.Read<ProductRecord>(stream)
                           .Validate(r => r.Price > 0, "Price must be strictly positive")
                           .ToList();

        // Assert
        Assert.Single(results); // Row 2 has an invalid price, so it's skipped by default
        Assert.Equal(1, results[0].Id);
        Assert.Equal("Apple", results[0].Name);
        Assert.Equal(1.50m, results[0].Price);
    }

    [Fact]
    public void Read_SkipRows_SkipsTitleAndReadsSuccessfully()
    {
        // Arrange
        using var stream = GenerateTestExcelStream(includeTitleRow: true);
        
        // Act
        var results = Excel.Read<ProductRecord>(stream)
                           .SkipRows(1)
                           .Validate(r => r.Price > 0, "Price must be strictly positive")
                           .ToList();

        // Assert
        Assert.Single(results);
        Assert.Equal("Apple", results[0].Name);
    }

    [Fact]
    public void Read_OnError_CatchesTypeErrors()
    {
        // Arrange
        using var stream = GenerateTestExcelStream();
        var errors = new List<ExcelParseError>();

        // Act
        var results = Excel.Read<ProductRecord>(stream)
                           .OnError(err => errors.Add(err))
                           .ToList();

        // Assert
        Assert.Single(errors);
        Assert.Equal("Price", errors[0].ColumnName);
        Assert.Equal("decimal", errors[0].ExpectedType);
    }

    [Fact]
    public async Task Write_GeneratesValidExcelAndReadsBack()
    {
        // Arrange
        var data = new List<ProductRecord>
        {
            new ProductRecord { Id = 10, Name = "Cherry", Price = 5.0m, AddedDate = new DateTime(2026, 5, 5) }
        };

        using var ms = new MemoryStream();
        
        // Act
        Excel.Write(data).ToSheet("MySheet").ToStream(ms);
        ms.Position = 0;

        var readBack = await Excel.Read<ProductRecord>(ms)
                                  .FromSheet("MySheet")
                                  .ToListAsync();

        // Assert
        Assert.Single(readBack);
        Assert.Equal("Cherry", readBack[0].Name);
        Assert.Equal(5.0m, readBack[0].Price);
        Assert.Equal(new DateTime(2026, 5, 5), readBack[0].AddedDate);
    }
}