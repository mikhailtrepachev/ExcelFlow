using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;

namespace ExcelFlow.SourceGenerators;

[Generator(LanguageNames.CSharp)]
public class MapperGenerator : IIncrementalGenerator
{
    private const string SerializableAttributeMetadataName = "ExcelFlow.ExcelFlowSerializableAttribute";
    private const string ColumnAttributeFullName = "ExcelFlow.ExcelColumnAttribute";

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        IncrementalValuesProvider<MapperResult> targets = context.SyntaxProvider
            .ForAttributeWithMetadataName(
                SerializableAttributeMetadataName,
                predicate: static (node, _) => node is ClassDeclarationSyntax,
                transform: static (ctx, ct) => BuildModel(ctx, ct));

        context.RegisterSourceOutput(targets, static (spc, result) => Emit(spc, result));
    }

    /// <summary>
    /// Extracts an equatable model from the semantic context. Only value-equatable data may
    /// leave this method (no ISymbol / SyntaxNode), otherwise the incremental cache is defeated.
    /// </summary>
    private static MapperResult BuildModel(GeneratorAttributeSyntaxContext ctx, CancellationToken cancellationToken)
    {
        if (ctx.TargetSymbol is not INamedTypeSymbol symbol)
            return new MapperResult(null, EquatableArray<DiagnosticInfo>.Empty);

        var classDecl = (ClassDeclarationSyntax)ctx.TargetNode;
        var diagnostics = new List<DiagnosticInfo>();
        bool fatal = false;

        if (symbol.IsGenericType)
        {
            diagnostics.Add(new DiagnosticInfo("EXF003", LocationInfo.From(classDecl), new EquatableArray<string>(new[] { symbol.Name })));
            fatal = true;
        }

        if (!classDecl.Modifiers.Any(SyntaxKind.PartialKeyword))
        {
            diagnostics.Add(new DiagnosticInfo("EXF001", LocationInfo.From(classDecl), new EquatableArray<string>(new[] { symbol.Name })));
            fatal = true;
        }

        // Enclosing type declarations must also be partial for the generated nesting to compile
        for (SyntaxNode? parent = classDecl.Parent; parent is not null; parent = parent.Parent)
        {
            if (parent is TypeDeclarationSyntax enclosingDecl && !enclosingDecl.Modifiers.Any(SyntaxKind.PartialKeyword))
            {
                diagnostics.Add(new DiagnosticInfo("EXF001", LocationInfo.From(enclosingDecl), new EquatableArray<string>(new[] { enclosingDecl.Identifier.Text })));
                fatal = true;
            }
        }

        var enclosingClasses = new List<string>();

        for (INamedTypeSymbol? current = symbol.ContainingType; current is not null; current = current.ContainingType)
        {
            if (current.IsGenericType)
            {
                diagnostics.Add(new DiagnosticInfo("EXF003", LocationInfo.From(classDecl), new EquatableArray<string>(new[] { current.Name })));
                fatal = true;
            }

            enclosingClasses.Insert(0, current.Name);
        }

        if (fatal)
            return new MapperResult(null, new EquatableArray<DiagnosticInfo>(diagnostics.ToArray()));

        var properties = new List<PropertyModel>();

        foreach (IPropertySymbol prop in symbol.GetMembers().OfType<IPropertySymbol>())
        {
            if (prop.IsStatic || prop.IsIndexer)
                continue;

            if (prop.DeclaredAccessibility != Accessibility.Public)
                continue;

            // Reading needs a public settable (non-init) property; writing needs a public getter.
            // Computed/read-only properties are intentionally skipped.
            if (prop.GetMethod is not { DeclaredAccessibility: Accessibility.Public })
                continue;

            if (prop.SetMethod is not { DeclaredAccessibility: Accessibility.Public, IsInitOnly: false })
                continue;

            AttributeData? columnAttr = prop.GetAttributes()
                .FirstOrDefault(a => a.AttributeClass?.ToDisplayString() == ColumnAttributeFullName);

            string columnName = prop.Name;
            int explicitIndex = -1;
            bool isRequired = false;

            if (columnAttr is not null)
            {
                if (columnAttr.ConstructorArguments.Length > 0 && columnAttr.ConstructorArguments[0].Value is string name && name.Length > 0)
                    columnName = name;

                foreach (KeyValuePair<string, TypedConstant> namedArg in columnAttr.NamedArguments)
                {
                    if (namedArg.Key == "Index" && namedArg.Value.Value is int idx)
                        explicitIndex = idx;

                    if (namedArg.Key == "IsRequired" && namedArg.Value.Value is bool req)
                        isRequired = req;
                }
            }

            ITypeSymbol type = prop.Type;
            bool isNullableValueType = type.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T;
            ITypeSymbol underlying = isNullableValueType ? ((INamedTypeSymbol)type).TypeArguments[0] : type;

            CellKind kind = Classify(underlying);

            if (kind == CellKind.Other)
            {
                SyntaxNode? propNode = prop.DeclaringSyntaxReferences.FirstOrDefault()?.GetSyntax(cancellationToken);
                diagnostics.Add(new DiagnosticInfo("EXF002", LocationInfo.From(propNode), new EquatableArray<string>(new[] { prop.Name, type.ToDisplayString() })));
            }

            properties.Add(new PropertyModel(
                Name: prop.Name,
                UnderlyingTypeFq: underlying.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat),
                DisplayType: type.ToDisplayString(),
                Kind: kind,
                IsNullableValueType: isNullableValueType,
                IsReferenceType: type.IsReferenceType,
                ColumnName: columnName,
                ExplicitIndex: explicitIndex,
                IsRequired: isRequired));
        }

        var model = new MapperModel(
            Namespace: symbol.ContainingNamespace.IsGlobalNamespace ? null : symbol.ContainingNamespace.ToDisplayString(),
            ClassName: symbol.Name,
            EnclosingClasses: new EquatableArray<string>(enclosingClasses.ToArray()),
            Properties: new EquatableArray<PropertyModel>(properties.ToArray()));

        return new MapperResult(model, new EquatableArray<DiagnosticInfo>(diagnostics.ToArray()));
    }

    private static CellKind Classify(ITypeSymbol type)
    {
        switch (type.SpecialType)
        {
            case SpecialType.System_String:
                return CellKind.String;
            case SpecialType.System_Boolean:
                return CellKind.Bool;
            case SpecialType.System_Byte:
            case SpecialType.System_SByte:
            case SpecialType.System_Int16:
            case SpecialType.System_UInt16:
            case SpecialType.System_Int32:
            case SpecialType.System_UInt32:
            case SpecialType.System_Int64:
            case SpecialType.System_UInt64:
            case SpecialType.System_Single:
            case SpecialType.System_Double:
            case SpecialType.System_Decimal:
                return CellKind.Number;
            case SpecialType.System_DateTime:
                return CellKind.DateTime;
        }

        if (type.TypeKind == TypeKind.Enum)
            return CellKind.Enum;

        if (type.ContainingNamespace is { Name: "System" } ns && ns.ContainingNamespace.IsGlobalNamespace && type.Name == "Guid")
            return CellKind.Guid;

        return CellKind.Other;
    }

    private static void Emit(SourceProductionContext context, MapperResult result)
    {
        foreach (DiagnosticInfo diagnostic in result.Diagnostics)
        {
            context.ReportDiagnostic(Diagnostic.Create(
                ExcelFlowDiagnostics.Get(diagnostic.Id),
                diagnostic.Location?.ToLocation() ?? Location.None,
                diagnostic.Args.Cast<object>().ToArray()));
        }

        if (result.Model is null)
            return;

        MapperModel model = result.Model;

        IEnumerable<string> hintParts = model.EnclosingClasses.Concat(new[] { model.ClassName });

        if (model.Namespace is not null)
            hintParts = new[] { model.Namespace }.Concat(hintParts);

        context.AddSource(string.Join(".", hintParts) + ".g.cs", SourceText.From(GenerateSource(model), Encoding.UTF8));
    }

    private static string GenerateSource(MapperModel model)
    {
        string className = model.ClassName;
        var sb = new StringBuilder();

        sb.AppendLine("// <auto-generated by ExcelFlow.SourceGenerators />");
        sb.AppendLine("#nullable enable");
        sb.AppendLine();
        sb.AppendLine("using XL = global::DocumentFormat.OpenXml.Spreadsheet;");
        sb.AppendLine();

        if (model.Namespace is not null)
        {
            sb.AppendLine($"namespace {model.Namespace}");
            sb.AppendLine("{");
        }

        foreach (string enclosing in model.EnclosingClasses)
        {
            sb.AppendLine($"    partial class {enclosing}");
            sb.AppendLine("    {");
        }

        sb.AppendLine($"    partial class {className} : global::ExcelFlow.IExcelFlowSerializable<{className}>");
        sb.AppendLine("    {");

        WriteGetDefinitions(sb, model);
        sb.AppendLine();
        WriteInitializeIndexMap(sb, model);
        sb.AppendLine();
        WriteParseRow(sb, model);
        sb.AppendLine();
        WriteWriteHeaders(sb, model);
        sb.AppendLine();
        WriteWriteRow(sb, model);

        sb.AppendLine("    }");

        foreach (string _ in model.EnclosingClasses)
            sb.AppendLine("    }");

        if (model.Namespace is not null)
            sb.AppendLine("}");

        return sb.ToString();
    }

    /// <summary>
    /// Escapes an arbitrary string as a valid C# string literal (quotes included).
    /// </summary>
    private static string L(string value) => SymbolDisplay.FormatLiteral(value, quote: true);

    private static void WriteGetDefinitions(StringBuilder sb, MapperModel model)
    {
        string className = model.ClassName;

        sb.AppendLine($"        public static global::System.Collections.Generic.IEnumerable<global::ExcelFlow.ExcelColumnDefinition<{className}>> GetDefinitions()");
        sb.AppendLine("        {");

        foreach (PropertyModel prop in model.Properties)
        {
            string propType = prop.IsNullableValueType ? prop.UnderlyingTypeFq + "?" : prop.UnderlyingTypeFq;

            sb.AppendLine($"            yield return new global::ExcelFlow.ExcelColumnDefinition<{className}>(");
            sb.AppendLine($"                {L(prop.ColumnName)},");
            sb.AppendLine($"                typeof({prop.UnderlyingTypeFq}),");
            sb.AppendLine($"                static (item, val) => item.{prop.Name} = val == null ? default({propType})! : ({propType})val,");
            sb.AppendLine($"                static (item) => item.{prop.Name});");
        }

        sb.AppendLine("            yield break;");
        sb.AppendLine("        }");
    }

    private static void WriteInitializeIndexMap(StringBuilder sb, MapperModel model)
    {
        sb.AppendLine("        public static int[] InitializeIndexMap(global::System.Collections.Generic.Dictionary<string, int> headerMap)");
        sb.AppendLine("        {");
        sb.AppendLine($"            int[] map = new int[{model.Properties.Count}];");

        for (int i = 0; i < model.Properties.Count; i++)
        {
            PropertyModel prop = model.Properties[i];

            if (prop.ExplicitIndex >= 0)
            {
                sb.AppendLine($"            map[{i}] = {prop.ExplicitIndex};");
                continue;
            }

            sb.AppendLine($"            if (headerMap.TryGetValue({L(prop.ColumnName)}, out int idx{i}))");
            sb.AppendLine($"                map[{i}] = idx{i};");
            sb.AppendLine("            else");

            if (prop.IsRequired)
                sb.AppendLine($"                throw new global::System.InvalidOperationException({L($"Required column '{prop.ColumnName}' is missing in the file.")});");
            else
                sb.AppendLine($"                map[{i}] = -1;");
        }

        sb.AppendLine("            return map;");
        sb.AppendLine("        }");
    }

    private static void WriteParseRow(StringBuilder sb, MapperModel model)
    {
        string className = model.ClassName;
        const string invariant = "global::System.Globalization.CultureInfo.InvariantCulture";

        sb.AppendLine($"        public static void ParseRow(global::ExcelDataReader.IExcelDataReader reader, int[] indexMap, out {className} item, global::System.Action<global::ExcelFlow.ExcelParseError>? onError, int rowNumber)");
        sb.AppendLine("        {");
        sb.AppendLine($"            item = new {className}();");

        for (int i = 0; i < model.Properties.Count; i++)
        {
            PropertyModel prop = model.Properties[i];
            string t = prop.UnderlyingTypeFq;
            string p = prop.Name;

            sb.AppendLine($"            if (indexMap[{i}] >= 0 && !reader.IsDBNull(indexMap[{i}]))");
            sb.AppendLine("            {");
            sb.AppendLine($"                object rawValue = reader.GetValue(indexMap[{i}]);");
            sb.AppendLine("                try");
            sb.AppendLine("                {");

            switch (prop.Kind)
            {
                case CellKind.String:
                    sb.AppendLine($"                    item.{p} = rawValue as string ?? rawValue.ToString()!;");
                    break;

                case CellKind.Number:
                    sb.AppendLine($"                    if (rawValue is double d) item.{p} = ({t})d;");
                    sb.AppendLine($"                    else if (rawValue is string s) item.{p} = {t}.Parse(s, {invariant});");
                    sb.AppendLine($"                    else item.{p} = ({t})global::System.Convert.ChangeType(rawValue, typeof({t}), {invariant});");
                    break;

                case CellKind.DateTime:
                    sb.AppendLine($"                    if (rawValue is global::System.DateTime dt) item.{p} = dt;");
                    sb.AppendLine($"                    else if (rawValue is double d) item.{p} = global::System.DateTime.FromOADate(d);");
                    sb.AppendLine($"                    else if (rawValue is string s) item.{p} = global::System.DateTime.Parse(s, {invariant});");
                    sb.AppendLine($"                    else item.{p} = (global::System.DateTime)global::System.Convert.ChangeType(rawValue, typeof(global::System.DateTime), {invariant});");
                    break;

                case CellKind.Bool:
                    sb.AppendLine($"                    if (rawValue is bool b) item.{p} = b;");
                    sb.AppendLine($"                    else if (rawValue is double d) item.{p} = d != 0;");
                    sb.AppendLine($"                    else if (rawValue is string s) item.{p} = bool.Parse(s);");
                    sb.AppendLine($"                    else item.{p} = (bool)global::System.Convert.ChangeType(rawValue, typeof(bool), {invariant});");
                    break;

                case CellKind.Guid:
                    sb.AppendLine($"                    if (rawValue is string s) item.{p} = global::System.Guid.Parse(s);");
                    sb.AppendLine($"                    else item.{p} = (global::System.Guid)rawValue;");
                    break;

                case CellKind.Enum:
                    sb.AppendLine($"                    if (rawValue is string s) item.{p} = global::System.Enum.Parse<{t}>(s, true);");
                    sb.AppendLine($"                    else if (rawValue is double d) item.{p} = ({t})(long)d;");
                    sb.AppendLine($"                    else item.{p} = ({t})rawValue;");
                    break;

                default:
                    sb.AppendLine($"                    item.{p} = ({t})global::System.Convert.ChangeType(rawValue, typeof({t}), {invariant});");
                    break;
            }

            sb.AppendLine("                }");
            sb.AppendLine("                catch (global::System.Exception)");
            sb.AppendLine("                {");
            sb.AppendLine($"                    onError?.Invoke(new global::ExcelFlow.ExcelParseError(rowNumber, {L(prop.Name)}, rawValue?.ToString() ?? \"\", {L(prop.DisplayType)}));");
            sb.AppendLine("                }");
            sb.AppendLine("            }");
        }

        sb.AppendLine("        }");
    }

    private static void WriteWriteHeaders(StringBuilder sb, MapperModel model)
    {
        sb.AppendLine("        public static void WriteHeaders(global::DocumentFormat.OpenXml.OpenXmlWriter writer)");
        sb.AppendLine("        {");
        sb.AppendLine("            writer.WriteStartElement(new XL.Row());");

        foreach (PropertyModel prop in model.Properties)
        {
            sb.AppendLine($"            writer.WriteElement(new XL.Cell {{ CellValue = new XL.CellValue({L(prop.ColumnName)}), DataType = XL.CellValues.String }});");
        }

        sb.AppendLine("            writer.WriteEndElement();");
        sb.AppendLine("        }");
    }

    private static void WriteWriteRow(StringBuilder sb, MapperModel model)
    {
        string className = model.ClassName;
        const string invariant = "global::System.Globalization.CultureInfo.InvariantCulture";

        sb.AppendLine($"        public static void WriteRow(global::DocumentFormat.OpenXml.OpenXmlWriter writer, {className} item)");
        sb.AppendLine("        {");
        sb.AppendLine("            writer.WriteStartElement(new XL.Row());");

        foreach (PropertyModel prop in model.Properties)
        {
            bool isNullable = prop.IsNullableValueType || prop.IsReferenceType;

            string access = prop.IsNullableValueType
                ? $"item.{prop.Name}!.Value"
                : prop.IsReferenceType ? $"item.{prop.Name}!" : $"item.{prop.Name}";

            string indent = "                ";

            sb.AppendLine("            {");

            if (isNullable)
            {
                sb.AppendLine($"                if (item.{prop.Name} is null)");
                sb.AppendLine("                {");
                sb.AppendLine("                    writer.WriteElement(new XL.Cell());");
                sb.AppendLine("                }");
                sb.AppendLine("                else");
                sb.AppendLine("                {");
                indent = "                    ";
            }

            sb.AppendLine($"{indent}var v = {access};");

            switch (prop.Kind)
            {
                case CellKind.Number:
                    sb.AppendLine($"{indent}writer.WriteElement(new XL.Cell {{ CellValue = new XL.CellValue(v.ToString({invariant})), DataType = XL.CellValues.Number }});");
                    break;

                case CellKind.DateTime:
                    // OADate number + date style so Excel renders it as a date, not a raw number.
                    // Style 1 = date only, style 2 = date + time (see ExcelExtensions.AddMinimalStylesheet).
                    sb.AppendLine($"{indent}writer.WriteElement(new XL.Cell {{ CellValue = new XL.CellValue(v.ToOADate().ToString({invariant})), DataType = XL.CellValues.Number, StyleIndex = v.TimeOfDay == global::System.TimeSpan.Zero ? 1U : 2U }});");
                    break;

                case CellKind.Bool:
                    sb.AppendLine($"{indent}writer.WriteElement(new XL.Cell {{ CellValue = new XL.CellValue(v ? \"1\" : \"0\"), DataType = XL.CellValues.Boolean }});");
                    break;

                case CellKind.String:
                    sb.AppendLine($"{indent}writer.WriteElement(new XL.Cell {{ CellValue = new XL.CellValue(v), DataType = XL.CellValues.String }});");
                    break;

                default:
                    sb.AppendLine($"{indent}writer.WriteElement(new XL.Cell {{ CellValue = new XL.CellValue(v.ToString() ?? \"\"), DataType = XL.CellValues.String }});");
                    break;
            }

            if (isNullable)
                sb.AppendLine("                }");

            sb.AppendLine("            }");
        }

        sb.AppendLine("            writer.WriteEndElement();");
        sb.AppendLine("        }");
    }
}
