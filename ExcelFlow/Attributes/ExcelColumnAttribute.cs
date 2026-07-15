namespace ExcelFlow;

/// <summary>
/// Maps a property to an Excel column by header name or by explicit index.
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
