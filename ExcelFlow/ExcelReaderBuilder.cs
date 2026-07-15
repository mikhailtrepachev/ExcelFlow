using System.Runtime.CompilerServices;

namespace ExcelFlow;

/// <summary>
/// ExcelReaderBuilder
/// </summary>
/// <typeparam name="T">Table scheme</typeparam>
public class ExcelReaderBuilder<T> where T : class, IExcelFlowSerializable<T>, new()
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
    /// If true, a caller-provided stream is left open after reading (the caller owns it).
    /// </summary>
    private readonly bool _leaveOpen;

    /// <summary>
    /// Sheet name - default is the first sheet
    /// </summary>
    private string? _sheetName = null;

    private Action<ExcelParseError>? _errorHandler;
    private int _skipRows = 0;

    private readonly List<(Func<T, bool> Predicate, string ErrorMessage)> _validationRules = new();

    /// <summary>
    /// Constructor for ExcelReaderBuilder
    /// </summary>
    /// <param name="filePath">File path</param>
    /// <param name="stream">Stream</param>
    /// <param name="leaveOpen">Whether a caller-provided stream should stay open after reading</param>
    internal ExcelReaderBuilder(string? filePath = null, Stream? stream = null, bool leaveOpen = true)
    {
        _filePath = filePath;
        _stream = stream;
        _leaveOpen = leaveOpen;
    }

    /// <summary>
    /// Allows you to select the worksheet from which the data will be retrieved
    /// </summary>
    /// <param name="sheetName">Sheet name</param>
    /// <returns>ExcelReaderBuilder</returns>
    public ExcelReaderBuilder<T> FromSheet(string? sheetName)
    {
        _sheetName = sheetName;
        return this;
    }

    /// <summary>
    /// Skips the given number of rows before the header row (e.g. title/description rows)
    /// </summary>
    /// <param name="count">Number of rows to skip; must be non-negative</param>
    /// <returns>ExcelReaderBuilder</returns>
    public ExcelReaderBuilder<T> SkipRows(int count)
    {
        if (count < 0)
            throw new ArgumentOutOfRangeException(nameof(count), count, "Row count to skip cannot be negative");

        _skipRows = count;
        return this;
    }

    /// <summary>
    /// Registers an error handler. Can be called multiple times; all handlers are invoked.
    /// </summary>
    /// <param name="handler">Error handler</param>
    /// <returns>Excel Reader builder</returns>
    public ExcelReaderBuilder<T> OnError(Action<ExcelParseError> handler)
    {
        if (handler is null)
            throw new ArgumentNullException(nameof(handler));

        _errorHandler += handler;
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
            : new ExcelContext(_stream!, _leaveOpen);

        foreach (T item in context.Worksheet<T>(_sheetName, _skipRows, _errorHandler, _validationRules))
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
            : new ExcelContext(_stream!, _leaveOpen);

        IAsyncEnumerable<T> asyncStream = context.WorksheetAsync<T>(_sheetName, _skipRows, _errorHandler, _validationRules, cancellationToken);

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

    /// <summary>
    /// Reserved for future use.
    /// </summary>
    [Obsolete("WithMapping has no effect in the current version and will be redesigned in v2.0. Column mapping is resolved at compile time by the source generator.")]
    public ExcelReaderBuilder<T> WithMapping(IEnumerable<ExcelColumnDefinition<T>> definitions)
    {
        return this;
    }

    /// <summary>
    /// Adds validation (without reflection)
    /// </summary>
    /// <param name="predicate"></param>
    /// <param name="errorMessage"></param>
    /// <returns></returns>
    /// <exception cref="ArgumentNullException"></exception>
    public ExcelReaderBuilder<T> Validate(Func<T, bool> predicate, string errorMessage)
    {
        if (predicate is null)
            throw new ArgumentNullException(nameof(predicate));

        _validationRules.Add((predicate, errorMessage));
        return this;
    }
}
