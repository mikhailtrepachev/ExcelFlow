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
    /// Will read the stream and return a builder.
    /// The stream is left open after reading: the caller owns it and is responsible for disposing it.
    /// </summary>
    ExcelReaderBuilder<T> Read<T>(Stream stream) where T : class, IExcelFlowSerializable<T>, new();

    /// <summary>
    /// Will read the stream and return a builder. Pass leaveOpen: false to let ExcelFlow dispose the stream.
    /// </summary>
    ExcelReaderBuilder<T> Read<T>(Stream stream, bool leaveOpen) where T : class, IExcelFlowSerializable<T>, new()
        => Excel.Read<T>(stream, leaveOpen);

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

    public ExcelReaderBuilder<T> Read<T>(Stream stream, bool leaveOpen) where T : class, IExcelFlowSerializable<T>, new()
        => Excel.Read<T>(stream, leaveOpen);

    public ExcelWriterBuilder<T> Write<T>(IEnumerable<T> data) where T : class, IExcelFlowSerializable<T>, new()
        => Excel.Write<T>(data);
}
