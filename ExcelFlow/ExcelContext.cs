using ExcelDataReader;
using System.Runtime.CompilerServices;

namespace ExcelFlow;

internal class ExcelContext : IDisposable
{
    private readonly Stream _fileStream;

    private readonly IExcelDataReader _reader;

    private readonly bool _leaveOpen;

    static ExcelContext()
    {
        // ExcelDataReader needs legacy code page encodings (e.g. for .xls); register them once.
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
    }

    /// <summary>
    /// Creates a new ExcelContext with the given file path
    /// </summary>
    /// <param name="filePath">Path to the file</param>
    public ExcelContext(string filePath)
    {
        // FileShare.Read lets the user keep the file open in Excel while we read it.
        _fileStream = new FileStream(
            filePath, FileMode.Open, FileAccess.Read, FileShare.Read, 4096, FileOptions.SequentialScan);

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
    /// <param name="leaveOpen">If true, the stream stays open (and owned by the caller) after disposing the context</param>
    /// <exception cref="ArgumentNullException"></exception>
    public ExcelContext(Stream stream, bool leaveOpen = false)
    {
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

        // 1-based sheet row number of the header row (skipped rows come before it)
        int rowNumber = skipRows + 1;

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

        // 1-based sheet row number of the header row (skipped rows come before it)
        int rowNumber = skipRows + 1;

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
            // A header cell can be empty (null) when data rows are wider than the header row
            string? colName = _reader.GetValue(i)?.ToString();

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
