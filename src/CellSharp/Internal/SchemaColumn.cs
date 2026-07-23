using System.Reflection;
using System.Collections.ObjectModel;

namespace CellSharp.Internal;

internal sealed class SchemaColumn
{
    internal SchemaColumn(
        PropertyInfo property,
        string header,
        bool isRequired,
        bool isIgnored,
        IEnumerable<ValidationRule> validations,
        DeclarativeValidationRule? declarativeValidation,
        ValueConverterDefinition? converter,
        string? format,
        double? width,
        ExcelHorizontalAlignment? alignment,
        ExcelVerticalAlignment? verticalAlignment,
        ExcelHorizontalAlignment? headerAlignment,
        ExcelVerticalAlignment? headerVerticalAlignment,
        int? sourceColumnNumber,
        string? sourceHeader,
        Func<ExcelFormulaContext, string>? formula)
    {
        Property = property;
        Header = header;
        IsRequired = isRequired;
        IsIgnored = isIgnored;
        Validations = new ReadOnlyCollection<ValidationRule>(validations.ToArray());
        DeclarativeValidation = declarativeValidation;
        Converter = converter;
        Format = format;
        Width = width;
        Alignment = alignment;
        VerticalAlignment = verticalAlignment;
        HeaderAlignment = headerAlignment;
        HeaderVerticalAlignment = headerVerticalAlignment;
        SourceColumnNumber = sourceColumnNumber;
        SourceHeader = sourceHeader;
        Formula = formula;
    }

    internal PropertyInfo Property { get; }

    internal string Header { get; }

    internal bool IsRequired { get; }

    internal bool IsIgnored { get; }

    internal IReadOnlyList<ValidationRule> Validations { get; }

    internal DeclarativeValidationRule? DeclarativeValidation { get; }

    internal ValueConverterDefinition? Converter { get; }

    internal string? Format { get; }

    internal double? Width { get; }

    internal ExcelHorizontalAlignment? Alignment { get; }

    internal ExcelVerticalAlignment? VerticalAlignment { get; }

    internal ExcelHorizontalAlignment? HeaderAlignment { get; }

    internal ExcelVerticalAlignment? HeaderVerticalAlignment { get; }

    internal int? SourceColumnNumber { get; }

    internal string? SourceHeader { get; }

    internal Func<ExcelFormulaContext, string>? Formula { get; }
}
