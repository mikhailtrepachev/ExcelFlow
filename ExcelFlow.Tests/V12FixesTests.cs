using ClosedXML.Excel;

namespace ExcelFlow.Tests;

public enum OrderStatus
{
    Pending,
    Shipped,
    Delivered
}

// Exercises Guid, enum, nullable value types and quotes in column names (generator escaping)
[ExcelFlowSerializable]
public partial class OrderRecord
{
    [ExcelColumn("Order ID")]
    public Guid OrderId { get; set; }

    [ExcelColumn("Status")]
    public OrderStatus Status { get; set; }

    [ExcelColumn("Discount")]
    public decimal? Discount { get; set; }

    [ExcelColumn("Quote \"Q1\"")]
    public string? Quote { get; set; }

    [ExcelColumn("Created")]
    public DateTime Created { get; set; }
}

[ExcelFlowSerializable]
public partial class PricedRecord
{
    [ExcelColumn("ID")]
    public int Id { get; set; }

    [ExcelColumn("Price")]
    public decimal Price { get; set; }
}

[ExcelFlowSerializable]
public partial class RequiredColumnRecord
{
    [ExcelColumn("Nonexistent Column", IsRequired = true)]
    public string? Value { get; set; }
}

public class V12FixesTests
{
    private static MemoryStream BuildSheet(Action<IXLWorksheet> fill, string sheetName = "Data")
    {
        var ms = new MemoryStream();

        using (var wb = new XLWorkbook())
        {
            fill(wb.Worksheets.Add(sheetName));
            wb.SaveAs(ms);
        }

        ms.Position = 0;
        return ms;
    }

    [Fact]
    public void Read_EmptyHeaderCell_DoesNotThrow()
    {
        // Data row is wider than the header row -> header cell 5 is null
        using var stream = BuildSheet(ws =>
        {
            ws.Cell(1, 1).Value = "ID";
            ws.Cell(1, 2).Value = "Price";
            ws.Cell(2, 1).Value = 1;
            ws.Cell(2, 2).Value = 2.5;
            ws.Cell(2, 5).Value = "orphan value without a header";
        });

        var results = Excel.Read<PricedRecord>(stream).ToList();

        Assert.Single(results);
        Assert.Equal(1, results[0].Id);
        Assert.Equal(2.5m, results[0].Price);
    }

    [Fact]
    public void Read_OnError_ReportsRealSheetRowNumber_WithSkipRows()
    {
        // Sheet layout: row 1 = title, row 2 = header, row 3 = valid, row 4 = invalid price
        using var stream = BuildSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Quarterly report";
            ws.Cell(2, 1).Value = "ID";
            ws.Cell(2, 2).Value = "Price";
            ws.Cell(3, 1).Value = 1;
            ws.Cell(3, 2).Value = 10.0;
            ws.Cell(4, 1).Value = 2;
            ws.Cell(4, 2).Value = "Free"; // type error here
        });

        var errors = new List<ExcelParseError>();

        Excel.Read<PricedRecord>(stream)
             .SkipRows(1)
             .OnError(errors.Add)
             .ToList();

        var error = Assert.Single(errors);
        Assert.Equal("Price", error.ColumnName);
        Assert.Equal(4, error.RowNumber); // real 1-based sheet row, not relative index
    }

    [Fact]
    public void Read_Stream_IsLeftOpenByDefault()
    {
        var stream = BuildSheet(ws =>
        {
            ws.Cell(1, 1).Value = "ID";
            ws.Cell(1, 2).Value = "Price";
            ws.Cell(2, 1).Value = 1;
            ws.Cell(2, 2).Value = 3.5;
        });

        Excel.Read<PricedRecord>(stream).ToList();

        Assert.True(stream.CanRead); // the caller owns the stream, it must survive reading
        stream.Dispose();
    }

    [Fact]
    public void Read_Stream_LeaveOpenFalse_DisposesStream()
    {
        var stream = BuildSheet(ws =>
        {
            ws.Cell(1, 1).Value = "ID";
            ws.Cell(1, 2).Value = "Price";
            ws.Cell(2, 1).Value = 1;
            ws.Cell(2, 2).Value = 3.5;
        });

        Excel.Read<PricedRecord>(stream, leaveOpen: false).ToList();

        Assert.False(stream.CanRead); // opted into library-owned disposal
    }

    [Fact]
    public void Read_OnError_MultipleHandlers_AllInvoked()
    {
        using var stream = BuildSheet(ws =>
        {
            ws.Cell(1, 1).Value = "ID";
            ws.Cell(1, 2).Value = "Price";
            ws.Cell(2, 1).Value = 1;
            ws.Cell(2, 2).Value = "not a number";
        });

        int first = 0, second = 0;

        Excel.Read<PricedRecord>(stream)
             .OnError(_ => first++)
             .OnError(_ => second++)
             .ToList();

        Assert.Equal(1, first);
        Assert.Equal(1, second);
    }

    [Fact]
    public void Read_RequiredColumnMissing_Throws()
    {
        using var stream = BuildSheet(ws =>
        {
            ws.Cell(1, 1).Value = "Some Other Column";
            ws.Cell(2, 1).Value = "value";
        });

        Assert.Throws<InvalidOperationException>(() => Excel.Read<RequiredColumnRecord>(stream).ToList());
    }

    [Fact]
    public void Write_DateTime_IsStyledAsDate()
    {
        var data = new List<OrderRecord>
        {
            new OrderRecord
            {
                OrderId = Guid.NewGuid(),
                Status = OrderStatus.Shipped,
                Created = new DateTime(2026, 5, 5) // date only
            },
            new OrderRecord
            {
                OrderId = Guid.NewGuid(),
                Status = OrderStatus.Pending,
                Created = new DateTime(2026, 5, 5, 14, 30, 0) // date + time
            }
        };

        using var ms = new MemoryStream();
        Excel.Write(data).ToSheet("Orders").ToStream(ms);
        ms.Position = 0;

        using var wb = new XLWorkbook(ms);
        var ws = wb.Worksheet("Orders");

        // ClosedXML detects DateTime only when the cell carries a date number format
        Assert.Equal(XLDataType.DateTime, ws.Cell(2, 5).DataType);
        Assert.Equal(new DateTime(2026, 5, 5), ws.Cell(2, 5).GetDateTime());
        Assert.Equal(XLDataType.DateTime, ws.Cell(3, 5).DataType);
        Assert.Equal(new DateTime(2026, 5, 5, 14, 30, 0), ws.Cell(3, 5).GetDateTime());
    }

    [Fact]
    public async Task RoundTrip_GuidEnumNullableAndQuotedColumn()
    {
        var id = Guid.NewGuid();

        var data = new List<OrderRecord>
        {
            new OrderRecord
            {
                OrderId = id,
                Status = OrderStatus.Delivered,
                Discount = 12.5m,
                Quote = "hello",
                Created = new DateTime(2026, 3, 1)
            },
            new OrderRecord
            {
                OrderId = Guid.NewGuid(),
                Status = OrderStatus.Pending,
                Discount = null, // empty cell must stay null
                Quote = null,
                Created = new DateTime(2026, 3, 2)
            }
        };

        using var ms = new MemoryStream();
        Excel.Write(data).ToSheet("Orders").ToStream(ms);
        ms.Position = 0;

        var readBack = await Excel.Read<OrderRecord>(ms).FromSheet("Orders").ToListAsync();

        Assert.Equal(2, readBack.Count);
        Assert.Equal(id, readBack[0].OrderId);
        Assert.Equal(OrderStatus.Delivered, readBack[0].Status);
        Assert.Equal(12.5m, readBack[0].Discount);
        Assert.Equal("hello", readBack[0].Quote);
        Assert.Equal(new DateTime(2026, 3, 1), readBack[0].Created);
        Assert.Null(readBack[1].Discount);
        Assert.Null(readBack[1].Quote);
    }
}
