namespace ExcelFlow;

/// <summary>
/// Represents a predefined mapping definition provided by a Source Generator.
/// It contains the expected column name and the compiled setter/getter delegates.
/// </summary>
public record ExcelColumnDefinition<T>(
    string ColumnName,
    Type PropertyType,
    Action<T, object?>? Setter,
    Func<T, object?>? Getter);
