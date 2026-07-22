using System.Reflection;

namespace CellSharp.Internal;

internal sealed class ImportType<T>
{
    private ImportType(Func<T> create, IReadOnlyList<ImportProperty> properties)
    {
        Create = create;
        Properties = properties;
    }

    internal Func<T> Create { get; }

    internal IReadOnlyList<ImportProperty> Properties { get; }

    internal static ImportType<T> CreateMetadata(ExcelSchema<T>? schema, ExcelSchemaOverlay<T>? overlay = null)
    {
        var type = typeof(T);
        if (type.IsAbstract || (!type.IsValueType && type.GetConstructor(Type.EmptyTypes) is null))
        {
            throw new InvalidOperationException($"Type '{type.FullName}' must have a public parameterless constructor to be read.");
        }

        var properties = type
            .GetProperties(BindingFlags.Instance | BindingFlags.Public)
            .Where(property => property.SetMethod?.IsPublic == true && property.GetIndexParameters().Length == 0)
            .OrderBy(property => property.MetadataToken)
            .ToArray();

        ImportProperty[] importProperties;
        if (schema is null)
        {
            if (properties.Length == 0)
            {
                throw new InvalidOperationException($"Type '{type.FullName}' has no public writable properties to read.");
            }

            var unsupported = properties.FirstOrDefault(property => !CellValueConverter.Supports(property.PropertyType));
            if (unsupported is not null)
            {
                throw new NotSupportedException(
                    $"Property '{unsupported.Name}' on type '{type.FullName}' has unsupported type '{unsupported.PropertyType.FullName}'.");
            }

            importProperties = properties.Select(property => new ImportProperty(property, property.Name, true)).ToArray();
        }
        else
        {
            importProperties = RuntimeSchema<T>.Create(schema, overlay).Columns
                .Select(column => new ImportProperty(
                    column.Column.Property,
                    column.Header,
                    column.Column.IsRequired,
                    column.Column.Validations,
                    column.Column.DeclarativeValidation,
                    column.Column.Converter,
                    column.Column.SourceColumnNumber,
                    column.Column.SourceHeader))
                .ToArray();
        }

        return new ImportType<T>(
            () => (T)(Activator.CreateInstance(type) ?? throw new InvalidOperationException($"Type '{type.FullName}' could not be created.")),
            importProperties);
    }
}
