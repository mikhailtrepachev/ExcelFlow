using ExcelDataReader;
using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using System.Runtime.CompilerServices;

namespace ExcelFlow;

/// <summary>
/// ExcelColumnAttribute
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
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

/// <summary>
/// Represents a predefined mapping definition provided by a Source Generator or Reflection fallback.
/// It contains the expected column name and the compiled setter delegate.
/// </summary>
public record ExcelColumnDefinition<T>(
    string ColumnName,
    [property: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] Type PropertyType,
    Action<T, object?>? Setter,
    Func<T, object?>? Getter);

/// <summary>
/// Represents an active mapping between an actual Excel column index and a model's property setter.
/// </summary>
internal record ColumnMapEntry<T>(
    int Index,
    string ColumnName,
    [property: DynamicallyAccessedMembers(DynamicallyAccessedMemberTypes.PublicParameterlessConstructor)] Type PropertyType,
    Action<T, object> Setter);

internal record ExportColumnMap<T>(
    string ColumnName,
    Type PropertyType,
    Func<T, object> Getter);

internal class ExcelContext : IDisposable
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

    public IEnumerable<T> Worksheet<T>(IEnumerable<ExcelColumnDefinition<T>> columnDefinitions,
        string? sheetName = null, Action<ExcelParseError>? onError = null)
        where T : new()
    {
        List<ColumnMapEntry<T>>? columnMap = PrepareColumnMap(columnDefinitions, sheetName);

        if (columnMap == null)
            yield break;

        int rowNumber = 1;

        while (_reader.Read())
        {
            rowNumber++;
            T item = new T();

            foreach (ColumnMapEntry<T> col in columnMap)
            {
                object? rawValue = _reader.GetValue(col.Index);
                (object? safeValue, bool isSuccess) = ExcelExtensions.SafeConvert(rawValue, col.PropertyType);

                if (!isSuccess)
                {
                    onError?.Invoke(new ExcelParseError(
                        RowNumber: rowNumber,
                        ColumnName: col.ColumnName,
                        RawValue: rawValue?.ToString() ?? string.Empty,
                        ExpectedType: col.PropertyType.Name
                    ));
                    continue;
                }

                col.Setter(item, safeValue ?? string.Empty);
            }

            yield return item;
        }
    }
    
    public async IAsyncEnumerable<T> WorksheetAsync<T>(IEnumerable<ExcelColumnDefinition<T>> columnDefinitions,
        string? sheetName = null,
        Action<ExcelParseError>? onError = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : new()
    {
        List<ColumnMapEntry<T>>? columnMap = PrepareColumnMap(columnDefinitions, sheetName);

        if (columnMap == null)
            yield break;
        
        int rowNumber = 1;

        while (_reader.Read())
        {
            cancellationToken.ThrowIfCancellationRequested();

            // Yield control back to the thread pool periodically to prevent thread starvation in async web apps
            if (rowNumber % 1000 == 0)
            {
                await Task.Yield();
            }
            
            rowNumber++;
            T item = new T();
            
            foreach (ColumnMapEntry<T> col in columnMap)
            {
                object? rawValue = _reader.GetValue(col.Index);
                (object? safeValue, bool isSuccess) = ExcelExtensions.SafeConvert(rawValue, col.PropertyType);

                if (!isSuccess)
                {
                    onError?.Invoke(new ExcelParseError(
                        RowNumber: rowNumber,
                        ColumnName: col.ColumnName, 
                        RawValue: rawValue?.ToString() ?? string.Empty,
                        ExpectedType: col.PropertyType.Name
                    ));
                    continue;
                }

                col.Setter(item, safeValue ?? string.Empty);
            }
            
            yield return item;
        }
    }

    /// <summary>
    /// Finds the requested worksheet, reads the header row, and matches it against the provided column definitions.
    /// </summary>
    /// <param name="columnDefinitions">The expected columns and their setters.</param>
    /// <param name="sheetName">The target sheet name (optional).</param>
    /// <returns>A list of active column mappings, or null if the sheet is empty.</returns>
    /// <exception cref="Exception">Sheet with this sheetName is not found.</exception>
    private List<ColumnMapEntry<T>>? PrepareColumnMap<T>(IEnumerable<ExcelColumnDefinition<T>> columnDefinitions,
        string? sheetName)
    {
        bool sheetFound = false;

        // Advance through result sets (sheets) until we find the requested one
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

        // Read the first row (headers). If false, the sheet is completely empty.
        if (!_reader.Read())
            return null;
        
        Dictionary<string, int> headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < _reader.FieldCount; i++)
        {
            string? colName = _reader.GetValue(i)?.ToString();

            if (!string.IsNullOrEmpty(colName))
            {
                headerMap[colName] = i;
            }
        }

        List<ColumnMapEntry<T>> columnMap = new List<ColumnMapEntry<T>>();

        foreach (ExcelColumnDefinition<T> definition in columnDefinitions)
        {
            if (definition.Setter != null && headerMap.TryGetValue(definition.ColumnName, out int index))
            {
                columnMap.Add(new ColumnMapEntry<T>(
                    index,
                    definition.ColumnName,
                    definition.PropertyType,
                    definition.Setter
                ));
            }
        }

        return columnMap;
    }

    /// <summary>
    /// Disposes the underlying reader and stream resources.
    /// </summary>
    public void Dispose()
    {
        _reader.Dispose();
        
        if (!_leaveOpen)
            _fileStream.Dispose();
    }
}