using System.Reflection;
using ExcelDataReader;

namespace ExcelFlow;

/// <summary>
/// ExcelColumnAttribute
/// </summary>
public class ExcelColumnAttribute : Attribute
{
    /// <summary>
    /// Name
    /// </summary>
    public string Name { get; }
    
    /// <summary>
    /// Constructor 
    /// </summary>
    /// <param name="name">Name</param>
    public ExcelColumnAttribute(string name) => Name = name;
}

public record ColumnMapEntry<T>(
    int Index,
    string ColumnName,
    Type PropertyType,
    Action<T, object> Setter);

public record ExportColumnMap<T>(
    string ColumnName,
    Type PropertyType,
    Func<T, object> Getter);

public class ExcelContext : IDisposable
{
    private readonly Stream _fileStream;

    private readonly IExcelDataReader _reader;

    private readonly bool _leaveOpen;

    /// <summary>
    /// Creates a new ExcelContext with the given file path 
    /// </summary>
    /// <param name="filePath">Path to the file</param>
    public ExcelContext(string filePath)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        
        _fileStream = File.Open(filePath, FileMode.Open, FileAccess.Read);
        
        _leaveOpen = false;
        
        _reader = ExcelReaderFactory.CreateReader(_fileStream, new ExcelReaderConfiguration
        {
            LeaveOpen = _leaveOpen
        });
    }

    /// <summary>
    /// Creates a new ExcelContext with the given stream
    /// </summary>
    /// <param name="stream">Stream</param>
    /// <param name="leaveOpen"></param>
    /// <exception cref="ArgumentNullException"></exception>
    public ExcelContext(Stream stream, bool leaveOpen = false)
    {
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        
        _fileStream = stream ?? throw new ArgumentNullException(nameof(stream));
        _leaveOpen = leaveOpen;
        
        _reader = ExcelReaderFactory.CreateReader(_fileStream, new ExcelReaderConfiguration
        {
            LeaveOpen = leaveOpen
        });
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sheetName"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    /// <exception cref="Exception"></exception>
    /// <exception cref="NotImplementedException"></exception>
    public IEnumerable<T> Worksheet<T>(string sheetName = null, Action<ExcelParseError> onError = null) where T : new()
    {
        bool sheetFound = false;

        do
        {
            if (sheetName == null || _reader.Name == sheetName)
            {
                sheetFound = true;
                break;
            }
        } while (_reader.NextResult());
        
        if (!sheetFound)
            throw new Exception($"Sheet {sheetName} not found");
        
        // Read header
        
        if (!_reader.Read())
            yield break;
        
        Dictionary<string, int> headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < _reader.FieldCount; i++)
        {
            string? colName = _reader.GetValue(i)?.ToString();

            if (!string.IsNullOrEmpty(colName))
            {
                headerMap[colName] = i;
            }
        }

        PropertyInfo[] properties = typeof(T).GetProperties();
        List<ColumnMapEntry<T>> columnMap = new List<ColumnMapEntry<T>>();

        foreach (PropertyInfo property in properties)
        {
            if (!property.CanWrite) 
                continue;
            
            ExcelColumnAttribute? attribute = property.GetCustomAttribute<ExcelColumnAttribute>();
            string expectedName = attribute?.Name ?? property.Name;

            if (headerMap.TryGetValue(expectedName, out int index))
            {
                Action<T, object> setter = ExpressionCompiler.CompileSetter<T>(property);
        
                columnMap.Add(new ColumnMapEntry<T>(
                    index, 
                    expectedName, 
                    property.PropertyType, 
                    setter
                ));
            }
        }
        
        int rowNumber = 1;

        while (_reader.Read())
        {
            rowNumber++;
            
            T? item = new T();

            foreach (ColumnMapEntry<T> col in columnMap)
            {
                object? rawValue = _reader.GetValue(col.Index);
        
                (object? safeValue, bool isSuccess) = ExcelExtensions.SafeConvert(rawValue, col.PropertyType);

                if (!isSuccess && onError != null)
                {
                    onError.Invoke(new ExcelParseError(
                        RowNumber: rowNumber,
                        ColumnName: col.ColumnName, 
                        RawValue: rawValue?.ToString(),
                        ExpectedType: col.PropertyType.Name
                    ));
                }

                col.Setter(item, safeValue);
            }
            
            yield return item;
        }
    }
    
    public void Dispose()
    {
        _reader.Dispose();
        
        if (!_leaveOpen)
            _fileStream.Dispose();
    }
}