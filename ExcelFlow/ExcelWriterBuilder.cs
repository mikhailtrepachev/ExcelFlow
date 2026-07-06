namespace ExcelFlow;

public class ExcelWriterBuilder<T> where T : class, IExcelFlowSerializable<T>, new()
{
    private readonly IEnumerable<T> _data;
    
    private string _sheetName = "Sheet1";

    internal ExcelWriterBuilder(IEnumerable<T> data)
    {
        _data = data;
    }

    public ExcelWriterBuilder<T> ToSheet(string sheetName)
    {
        _sheetName = sheetName;
        return this;
    }

    public void ToFile(string filePath)
    {
        _data.ToExcel(filePath, _sheetName);
    }

    public void ToStream(Stream stream)
    {
        _data.ToExcel(stream, _sheetName);
    }
}