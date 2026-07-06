using System;

namespace ExcelFlow.Attributes;

/// <summary>
/// Specifies the numeric range constraints for the property.
/// </summary>
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