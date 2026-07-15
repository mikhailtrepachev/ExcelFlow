using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.Text;

namespace ExcelFlow.SourceGenerators;

/// <summary>
/// How a cell value is converted to/from the property type in generated code.
/// </summary>
internal enum CellKind
{
    String,
    Number,
    Bool,
    DateTime,
    Guid,
    Enum,

    /// <summary>Falls back to Convert.ChangeType at runtime; reported as EXF002.</summary>
    Other
}

/// <summary>
/// Equatable snapshot of a mapped property. Must contain only value-equatable data
/// (no ISymbol references) so the incremental pipeline can cache it.
/// </summary>
internal sealed record PropertyModel(
    string Name,
    string UnderlyingTypeFq,
    string DisplayType,
    CellKind Kind,
    bool IsNullableValueType,
    bool IsReferenceType,
    string ColumnName,
    int ExplicitIndex,
    bool IsRequired);

/// <summary>
/// Equatable snapshot of a [ExcelFlowSerializable] class.
/// </summary>
internal sealed record MapperModel(
    string? Namespace,
    string ClassName,
    EquatableArray<string> EnclosingClasses,
    EquatableArray<PropertyModel> Properties);

/// <summary>
/// Equatable stand-in for Location (Location itself must not be cached in the pipeline).
/// </summary>
internal sealed record LocationInfo(string FilePath, TextSpan TextSpan, LinePositionSpan LineSpan)
{
    public Location ToLocation() => Location.Create(FilePath, TextSpan, LineSpan);

    public static LocationInfo? From(SyntaxNode? node)
        => node is null
            ? null
            : new LocationInfo(node.SyntaxTree.FilePath, node.Span, node.GetLocation().GetLineSpan().Span);
}

/// <summary>
/// Equatable stand-in for Diagnostic.
/// </summary>
internal sealed record DiagnosticInfo(string Id, LocationInfo? Location, EquatableArray<string> Args);

/// <summary>
/// Transform output: the model to emit (null when a fatal diagnostic was reported) plus diagnostics.
/// </summary>
internal sealed record MapperResult(MapperModel? Model, EquatableArray<DiagnosticInfo> Diagnostics);

internal static class ExcelFlowDiagnostics
{
    private const string Category = "ExcelFlow.SourceGenerator";

    public static readonly DiagnosticDescriptor ClassMustBePartial = new DiagnosticDescriptor(
        id: "EXF001",
        title: "ExcelFlow serializable class must be partial",
        messageFormat: "Type '{0}' is marked with [ExcelFlowSerializable] (or encloses such a class) and must be declared 'partial' so the source generator can extend it",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor UnsupportedPropertyType = new DiagnosticDescriptor(
        id: "EXF002",
        title: "Property type is not natively supported by ExcelFlow",
        messageFormat: "Property '{0}' has type '{1}' which is not natively supported by ExcelFlow. Reading falls back to Convert.ChangeType and may fail at runtime. Natively supported types: string, bool, numeric primitives, DateTime, Guid and enums (plus their nullable variants).",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Warning,
        isEnabledByDefault: true);

    public static readonly DiagnosticDescriptor GenericClassNotSupported = new DiagnosticDescriptor(
        id: "EXF003",
        title: "Generic classes are not supported",
        messageFormat: "Type '{0}' is generic; [ExcelFlowSerializable] does not support generic classes",
        category: Category,
        defaultSeverity: DiagnosticSeverity.Error,
        isEnabledByDefault: true);

    public static DiagnosticDescriptor Get(string id) => id switch
    {
        "EXF001" => ClassMustBePartial,
        "EXF002" => UnsupportedPropertyType,
        "EXF003" => GenericClassNotSupported,
        _ => throw new System.ArgumentOutOfRangeException(nameof(id), id, "Unknown diagnostic id")
    };
}
