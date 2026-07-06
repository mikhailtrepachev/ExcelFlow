using System;

namespace ExcelFlow.Attributes;

/// <summary>
/// Specifies that a property must have a non-empty value.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public class ExcelRequiredAttribute : Attribute
{
    public string? ErrorMessage { get; set; }
}