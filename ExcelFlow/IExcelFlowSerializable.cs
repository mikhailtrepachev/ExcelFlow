using ExcelDataReader;

namespace ExcelFlow;

/// <summary>
/// Defines a contract for models that provide compile-time generated Excel column definitions.
/// Required for NativeAOT and Trimming compatibility.
/// </summary>
/// <typeparam name="T">The type of the model implementing this interface.</typeparam>
public interface IExcelFlowSerializable<T> where T : new()
{
    static abstract IEnumerable<ExcelColumnDefinition<T>> GetDefinitions();

    static abstract int[] InitializeIndexMap(System.Collections.Generic.Dictionary<string, int> headerMap);
    
    static abstract void ParseRow(IExcelDataReader reader, int[] indexMap, out T item, Action<ExcelParseError>? onError, int rowNumber);
    
    static abstract void WriteHeaders(DocumentFormat.OpenXml.OpenXmlWriter writer);
    
    static abstract void WriteRow(DocumentFormat.OpenXml.OpenXmlWriter writer, T item);
}
