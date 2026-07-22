using System.Reflection;

namespace CellSharp.Internal;

internal sealed class RuntimeColumnOverride
{
    internal RuntimeColumnOverride(string? header, bool? isEnabled)
    {
        Header = header;
        IsEnabled = isEnabled;
    }

    internal string? Header { get; }

    internal bool? IsEnabled { get; }
}

internal sealed class RuntimeSchema<T>
{
    private RuntimeSchema(IReadOnlyList<RuntimeSchemaColumn> columns)
    {
        Columns = columns;
    }

    internal IReadOnlyList<RuntimeSchemaColumn> Columns { get; }

    internal static RuntimeSchema<T> Create(ExcelSchema<T> schema, ExcelSchemaOverlay<T>? overlay)
    {
        var configured = overlay?.Columns ?? new Dictionary<PropertyInfo, RuntimeColumnOverride>();
        foreach (var pair in configured)
        {
            var column = schema.Columns.FirstOrDefault(candidate => candidate.Property == pair.Key);
            if (column is null)
            {
                throw new ArgumentException($"Property '{pair.Key.Name}' is not configured by this schema.", nameof(overlay));
            }

            if (column.IsIgnored)
            {
                throw new ArgumentException($"Ignored property '{pair.Key.Name}' cannot be configured by a runtime overlay.", nameof(overlay));
            }
        }

        var columns = schema.Columns
            .Where(column => !column.IsIgnored)
            .Where(column => !configured.TryGetValue(column.Property, out var setting) || setting.IsEnabled != false)
            .Select(column => new RuntimeSchemaColumn(
                column,
                configured.TryGetValue(column.Property, out var setting) && setting.Header is not null
                    ? setting.Header
                    : column.Header))
            .ToArray();

        if (columns.Length == 0)
        {
            throw new InvalidOperationException("A runtime schema overlay must leave at least one active column.");
        }

        var duplicateHeader = columns.GroupBy(column => column.Header, StringComparer.OrdinalIgnoreCase)
            .FirstOrDefault(group => group.Count() > 1);
        if (duplicateHeader is not null)
        {
            throw new ArgumentException($"Runtime header '{duplicateHeader.Key}' is mapped to more than one active property.", nameof(overlay));
        }

        foreach (var source in columns.Where(column => column.Column.SourceHeader is not null))
        {
            var conflict = columns.FirstOrDefault(other => other.Column.Property != source.Column.Property
                && string.Equals(other.Header, source.Column.SourceHeader, StringComparison.OrdinalIgnoreCase));
            if (conflict is not null)
            {
                throw new ArgumentException(
                    $"Runtime header '{conflict.Header}' for property '{conflict.Column.Property.Name}' conflicts with MapFromHeader on property '{source.Column.Property.Name}'.",
                    nameof(overlay));
            }
        }

        return new RuntimeSchema<T>(columns);
    }
}

internal sealed class RuntimeSchemaColumn
{
    internal RuntimeSchemaColumn(SchemaColumn column, string header)
    {
        Column = column;
        Header = header;
    }

    internal SchemaColumn Column { get; }

    internal string Header { get; }
}
