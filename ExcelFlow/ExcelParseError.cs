namespace ExcelFlow;

public record ExcelParseError(
    int RowNumber,
    string ColumnName,
    string RawValue,
    string ExpectedType);