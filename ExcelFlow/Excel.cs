namespace ExcelFlow;

/// <summary>
/// Entry point for ExcelFlow
/// </summary>
public static class Excel
{
    /// <summary>
    /// Will read the file and return a builder
    /// </summary>
    /// <param name="filePath">File path</param>
    /// <typeparam name="T">Table scheme</typeparam>
    /// <returns>Excel Reader builder</returns>
    public static ExcelReaderBuilder<T> Read<T>(string filePath) where T : class, IExcelFlowSerializable<T>, new() 
        => new ExcelReaderBuilder<T>(filePath);


    /// <summary>
    /// Will read the stream and return a builder
    /// </summary>
    /// <param name="stream">Stream</param>
    /// <typeparam name="T">Table scheme</typeparam>
    /// <returns>Excel Reader builder</returns>
    public static ExcelReaderBuilder<T> Read<T>(Stream stream) where T : class, IExcelFlowSerializable<T>, new() 
        => new ExcelReaderBuilder<T>(null, stream);


    /// <summary>
    /// Will write the data to the file
    /// </summary>
    /// <param name="data">File path</param>
    /// <typeparam name="T">Table scheme</typeparam>
    /// <returns>Excel Writer builder</returns>
    public static ExcelWriterBuilder<T> Write<T>(IEnumerable<T> data) where T : class, IExcelFlowSerializable<T>, new()
        => new ExcelWriterBuilder<T>(data);
}