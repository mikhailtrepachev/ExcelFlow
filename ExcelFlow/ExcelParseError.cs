namespace ExcelFlow;

/// <summary>
/// ExcelParseError indicates data that could not be parsed
/// </summary>
/// <param name="RowNumber">Row number</param>
/// <param name="ColumnName">Column name</param>
/// <param name="RawValue">Value of cell</param>
/// <param name="ExpectedType">The expected data type</param>
public record ExcelParseError(
    int RowNumber,
    string ColumnName,
    string RawValue,
    string ExpectedType);