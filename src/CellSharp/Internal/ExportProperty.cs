using System.Reflection;

namespace CellSharp.Internal;

internal sealed class ExportProperty
{
    private ExportProperty(
        PropertyInfo property,
        string header,
        string? format = null,
        double? width = null,
        ExcelHorizontalAlignment? alignment = null,
        DeclarativeValidationRule? declarativeValidation = null,
        ValueConverterDefinition? converter = null,
        Func<ExcelFormulaContext, string>? formula = null)
    {
        Property = property;
        Header = header;
        Format = format;
        Width = width;
        Alignment = alignment;
        DeclarativeValidation = declarativeValidation;
        Converter = converter;
        Formula = formula;
    }

    internal PropertyInfo Property { get; }

    internal string Header { get; }

    internal string? Format { get; }

    internal double? Width { get; }

    internal ExcelHorizontalAlignment? Alignment { get; }

    internal DeclarativeValidationRule? DeclarativeValidation { get; }

    internal ValueConverterDefinition? Converter { get; }

    internal Func<ExcelFormulaContext, string>? Formula { get; }

    internal Type CellType => Converter?.CellType ?? Property.PropertyType;

    internal object? GetValue<T>(T item) => Property.GetValue(item);

    internal object? ConvertValue(object? value) => value is null
        ? null
        : Converter is null
            ? value
            : Converter.Write(value);

    internal static IReadOnlyList<ExportProperty> For<T>(ExcelSchema<T>? schema, ExcelSchemaOverlay<T>? overlay = null)
    {
        if (schema is not null)
        {
            return RuntimeSchema<T>.Create(schema, overlay).Columns
                .Select(column => new ExportProperty(
                    column.Column.Property,
                    column.Header,
                    column.Column.Format,
                    column.Column.Width,
                    column.Column.Alignment,
                    column.Column.DeclarativeValidation,
                    column.Column.Converter,
                    column.Column.Formula))
                .ToArray();
        }

        if (overlay is not null)
        {
            throw new ArgumentException("A runtime schema overlay requires an explicit schema.", nameof(overlay));
        }

        var properties = typeof(T)
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(property => property.CanRead && property.GetIndexParameters().Length == 0)
            .OrderBy(property => property.MetadataToken)
            .ToArray();

        if (properties.Length == 0)
        {
            throw new InvalidOperationException($"Type '{typeof(T).FullName}' has no public readable properties to export.");
        }

        var unsupported = properties.FirstOrDefault(property => !CellValueConverter.Supports(property.PropertyType));
        if (unsupported is not null)
        {
            throw new NotSupportedException(
                $"Property '{unsupported.Name}' on type '{typeof(T).FullName}' has unsupported type '{unsupported.PropertyType.FullName}'.");
        }

        return properties.Select(property => new ExportProperty(property, property.Name)).ToArray();
    }
}
