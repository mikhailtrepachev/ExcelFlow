using ExcelDataReader;
using System.Diagnostics.CodeAnalysis;
using System.Runtime.CompilerServices;

namespace ExcelFlow;

/// <summary>
/// ExcelColumnAttribute
/// </summary>
[AttributeUsage(AttributeTargets.Property)]
public class ExcelColumnAttribute : Attribute
{
    /// <summary>
    /// Expected column name
    /// </summary>
    public string? Name { get; }

    /// <summary>
    /// Explicit column index (0-based). If set, ignores the Name matching.
    /// </summary>
    public int Index { get; set; } = -1;

    /// <summary>
    /// If true, parser will throw an exception if the column is missing in the file.
    /// </summary>
    public bool IsRequired { get; set; }
    
    /// <summary>
    /// Constructor with name
    /// </summary>
    /// <param name="name">Name</param>
    public ExcelColumnAttribute(string name) => Name = name;

    /// <summary>
    /// Parameterless constructor for Index-based mapping
    /// </summary>
    public ExcelColumnAttribute() => Name = null;
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

    public IEnumerable<T> Worksheet<T>(
        string? sheetName = null, int skipRows = 0, Action<ExcelParseError>? onError = null, List<(Func<T, bool> Predicate, string ErrorMessage)>? validationRules = null)
        where T : class, IExcelFlowSerializable<T>, new()
    {
        int[]? indexMap = PrepareIndexMap<T>(sheetName, skipRows);

        if (indexMap == null)
            yield break;

        int rowNumber = 1;

        while (_reader.Read())
        {
            rowNumber++;
            T.ParseRow(_reader, indexMap, out T item, onError, rowNumber);

            bool isValidRow = true;

            if (validationRules is not null && validationRules.Count > 0)
            {
                foreach ((Func<T, bool> Predicate, string ErrorMessage) rule in validationRules)
                {
                    if (!rule.Predicate(item))
                    {
                        onError?.Invoke(new ExcelParseError(
                            RowNumber: rowNumber,
                            ColumnName: "Validation",
                            RawValue: rule.ErrorMessage,
                            ExpectedType: "BusinessRule"));

                        isValidRow = false;
                        break;
                    }
                }
            }

            if (!isValidRow)
                continue;

            yield return item;
        }
    }
    
    public async IAsyncEnumerable<T> WorksheetAsync<T>(
        string? sheetName = null,
        int skipRows = 0, Action<ExcelParseError>? onError = null, List<(Func<T, bool> Predicate, string ErrorMessage)>? validationRules = null,
        [EnumeratorCancellation] CancellationToken cancellationToken = default) where T : class, IExcelFlowSerializable<T>, new()
    {
        int[]? indexMap = PrepareIndexMap<T>(sheetName, skipRows);

        if (indexMap == null)
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
            T.ParseRow(_reader, indexMap, out T item, onError, rowNumber);

            bool isValidRow = true;

            if (validationRules is not null && validationRules.Count > 0)
            {
                foreach ((Func<T, bool> Predicate, string ErrorMessage) rule in validationRules)
                {
                    if (!rule.Predicate(item))
                    {
                        onError?.Invoke(new ExcelParseError(
                            RowNumber: rowNumber,
                            ColumnName: "Validation",
                            RawValue: rule.ErrorMessage,
                            ExpectedType: "BusinessRule"));

                        isValidRow = false;
                        break;
                    }
                }
            }

            if (!isValidRow)
                continue;

            yield return item;
        }
    }

    /// <summary>
    /// Prepares the column map for the given sheet.
    /// </summary>
    /// <param name="sheetName"></param>
    /// <param name="skipRows"></param>
    /// <typeparam name="T"></typeparam>
    /// <returns></returns>
    /// <exception cref="InvalidOperationException"></exception>
    private int[]? PrepareIndexMap<T>(
        string? sheetName, int skipRows) where T : class, IExcelFlowSerializable<T>, new()
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
            throw new InvalidOperationException($"Sheet {sheetName ?? "default"} not found");

        for (int i = 0; i < skipRows; i++)
        {
            if (!_reader.Read())
                return null;
        }

        // Read the first row (headers). If false, the sheet is completely empty.
        if (!_reader.Read())
            return null;
        
        Dictionary<string, int> headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        for (int i = 0; i < _reader.FieldCount; i++)
        {
            string? colName = _reader.GetValue(i).ToString();

            if (!string.IsNullOrEmpty(colName))
            {
                headerMap[colName] = i;
            }
        }

        return T.InitializeIndexMap(headerMap);
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