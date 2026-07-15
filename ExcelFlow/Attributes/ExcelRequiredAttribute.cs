namespace ExcelFlow.Attributes;

/// <summary>
/// Specifies that a property must have a non-empty value.
/// </summary>
[Obsolete("This attribute is not enforced yet and currently has no effect. It will be wired into generated validation in v2.0. Use [ExcelColumn(IsRequired = true)] for required columns or .Validate(...) for value checks.")]
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class ExcelRequiredAttribute : Attribute
{
    public string? ErrorMessage { get; set; }
}