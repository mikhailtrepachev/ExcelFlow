namespace ExcelFlow;

/// <summary>
/// A Dependency-Injection friendly service to work with ExcelFlow.
/// Use this interface in your ASP.NET Core controllers and services instead of the static Excel class.
/// </summary>
public interface IExcelFlowService
{
    /// <summary>
    /// Will read the file and return a builder
    /// </summary>
    ExcelReaderBuilder<T> Read<T>(string filePath) where T : class, IExcelFlowSerializable<T>, new();

    /// <summary>
    /// Will read the stream and return a builder
    /// </summary>
    ExcelReaderBuilder<T> Read<T>(Stream stream) where T : class, IExcelFlowSerializable<T>, new();

    /// <summary>
    /// Will write the data to the file/stream builder
    /// </summary>
    ExcelWriterBuilder<T> Write<T>(IEnumerable<T> data) where T : class, IExcelFlowSerializable<T>, new();
}

/// <summary>
/// Default implementation of IExcelFlowService
/// </summary>
public class ExcelFlowService : IExcelFlowService
{
    public ExcelReaderBuilder<T> Read<T>(string filePath) where T : class, IExcelFlowSerializable<T>, new()
        => Excel.Read<T>(filePath);

    public ExcelReaderBuilder<T> Read<T>(Stream stream) where T : class, IExcelFlowSerializable<T>, new()
        => Excel.Read<T>(stream);

    public ExcelWriterBuilder<T> Write<T>(IEnumerable<T> data) where T : class, IExcelFlowSerializable<T>, new()
        => Excel.Write<T>(data);
}
