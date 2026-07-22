using System.Reflection;

namespace CellSharp.Internal;

internal sealed class ImportProperty
{
    internal ImportProperty(
        PropertyInfo property,
        string header,
        bool isRequired,
        IReadOnlyList<ValidationRule>? validations = null,
        DeclarativeValidationRule? declarativeValidation = null,
        ValueConverterDefinition? converter = null,
        int? sourceColumnNumber = null,
        string? sourceHeader = null)
    {
        Property = property;
        Header = header;
        IsRequired = isRequired;
        Validations = validations ?? Array.Empty<ValidationRule>();
        DeclarativeValidation = declarativeValidation;
        Converter = converter;
        SourceColumnNumber = sourceColumnNumber;
        SourceHeader = sourceHeader;
    }

    internal PropertyInfo Property { get; }

    internal string Name => Property.Name;

    internal string Header { get; }

    internal bool IsRequired { get; }

    internal IReadOnlyList<ValidationRule> Validations { get; }

    internal DeclarativeValidationRule? DeclarativeValidation { get; }

    internal ValueConverterDefinition? Converter { get; }

    internal int? SourceColumnNumber { get; }

    internal string? SourceHeader { get; }

    internal Type Type => Property.PropertyType;

    internal void SetValue<T>(T item, object? value) => Property.SetValue(item, value);
}
