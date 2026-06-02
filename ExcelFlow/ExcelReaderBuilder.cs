using System.Runtime.CompilerServices;

namespace ExcelFlow;

/// <summary>
/// ExcelReaderBuilder
/// </summary>
/// <typeparam name="T">Table scheme</typeparam>
public class ExcelReaderBuilder<T> where T : class, new()
{
    /// <summary>
    /// File path
    /// </summary>
    private readonly string? _filePath;
    
    /// <summary>
    /// Stream
    /// </summary>
    private readonly Stream? _stream;
    
    /// <summary>
    /// Sheet name - default is "Sheet1"
    /// </summary>
    private string _sheetName = "Sheet1";
    
    private Action<ExcelParseError>? _errorHandler;

    /// <summary>
    /// Constructor for ExcelReaderBuilder
    /// </summary>
    /// <param name="filePath">File path</param>
    /// <param name="stream">Stream</param>
    internal ExcelReaderBuilder(string? filePath = null, Stream? stream = null)
    {
        _filePath = filePath;
        _stream = stream;
    }

    /// <summary>
    /// Allows you to select the worksheet from which the data will be retrieved
    /// </summary>
    /// <param name="sheetName">Sheet name</param>
    /// <returns>ExcelReaderBuilder</returns>
    public ExcelReaderBuilder<T> FromSheet(string sheetName)
    {
        _sheetName = sheetName;
        return this;
    }

    /// <summary>
    /// Error handler
    /// </summary>
    /// <param name="handler">Error handler</param>
    /// <returns>Excel Reader builder</returns>
    public ExcelReaderBuilder<T> OnError(Action<ExcelParseError> handler)
    {
        _errorHandler = handler;
        return this;
    }

    /// <summary>
    /// Returns the data from Excel table as an Enumerable
    /// </summary>
    /// <returns>IEnumerable</returns>
    public IEnumerable<T> AsEnumerable()
    {
        using ExcelContext context = _filePath is not null
            ? new ExcelContext(_filePath)
            : new ExcelContext(_stream!);

        foreach (T item in context.Worksheet<T>(_sheetName))
        {
            yield return item;
        }
    }
    
    /// <summary>
    /// Returns an asynchronous stream (<see cref="IAsyncEnumerable{T}"/>) for lazy reading of Excel rows.
    /// The file is read sequentially, row by row, which minimizes RAM consumption.
    /// </summary>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>IAsyncEnumerable</returns>
    public async IAsyncEnumerable<T> AsAsyncEnumerable(
        [EnumeratorCancellation] CancellationToken cancellationToken = default)
    {
        using ExcelContext context = _filePath != null 
            ? new ExcelContext(_filePath) 
            : new ExcelContext(_stream!);

        IAsyncEnumerable<T> asyncStream = context.WorksheetAsync<T>(_sheetName, _errorHandler, cancellationToken);

        await foreach (T item in asyncStream)
        {
            yield return item;
        }
    }

    /// <summary>
    /// Returns the data from Excel table as a List
    /// </summary>
    /// <returns>List</returns>
    public List<T> ToList() => AsEnumerable().ToList();

    /// <summary>
    /// Returns the data from Excel table as a List
    /// </summary>
    /// <param name="cancellationToken">Cancellation token</param>
    /// <returns>List</returns>
    public async Task<List<T>> ToListAsync(CancellationToken cancellationToken = default)
    {
        List<T> list = new List<T>();

        await foreach (T item in AsAsyncEnumerable(cancellationToken))
        {
            list.Add(item);
        }
        
        return list;
    }
}