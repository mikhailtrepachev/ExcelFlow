namespace ExcelFlow.Attributes;

/// <summary>
/// Specifies the numeric range constraints for the property.
/// </summary>
[Obsolete("This attribute is not enforced yet and currently has no effect. It will be wired into generated validation in v2.0. Use .Validate(...) for value checks.")]
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class ExcelRangeAttribute : Attribute
{
    public double Min { get; }
    public double Max { get; }
    public string? ErrorMessage { get; set; }

    public ExcelRangeAttribute(double min, double max)
    {
        Min = min;
        Max = max;
    }
}