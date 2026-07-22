using System.Reflection;
using CellSharp.Internal;

namespace CellSharp;

/// <summary>Provides the entry point for structured XLSX export.</summary>
public static class Excel
{
    /// <summary>Starts a typed XLSX schema definition.</summary>
    public static ExcelSchemaBuilder<T> Schema<T>() => new();

    /// <summary>
    /// Creates an immutable schema from public instance properties and their CellSharp attributes.
    /// Properties without CellSharp attributes retain the normal property-name defaults.
    /// </summary>
    public static ExcelSchema<T> SchemaFromAttributes<T>() => BuildAttributeSchema<T>(null);

    /// <summary>
    /// Creates an immutable schema from CellSharp attributes and applies fluent configuration over it.
    /// Fluent configuration for an attributed property replaces that property's attribute defaults.
    /// </summary>
    public static ExcelSchema<T> SchemaFromAttributes<T>(Action<ExcelSchemaBuilder<T>> configure)
    {
        if (configure is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(configure);
#else
            throw new ArgumentNullException(nameof(configure));
#endif
        }

        return BuildAttributeSchema(configure);
    }

    /// <summary>Builds an immutable per-operation overlay for a typed schema.</summary>
    public static ExcelSchemaOverlay<T> Overlay<T>(Action<ExcelSchemaOverlayBuilder<T>> configure)
    {
        if (configure is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(configure);
#else
            throw new ArgumentNullException(nameof(configure));
#endif
        }

        var builder = new ExcelSchemaOverlayBuilder<T>();
        configure(builder);
        return builder.Build();
    }

    /// <summary>Starts a typed multi-worksheet XLSX workbook definition.</summary>
    public static ExcelWorkbookBuilder Workbook() => new();

    /// <summary>Opens an XLSX workbook for multiple typed worksheet reads.</summary>
    public static ExcelWorkbookReader Open(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        return new ExcelWorkbookReader(path);
    }

    /// <summary>Opens an XLSX workbook with explicit package and per-sheet resource limits.</summary>
    public static ExcelWorkbookReader Open(string path, ExcelReadOptions options)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        ValidateReadOptions(options);
        return new ExcelWorkbookReader(path, options);
    }

    /// <summary>Opens a seekable readable XLSX stream for multiple typed reads without taking ownership of the stream.</summary>
    public static ExcelWorkbookReader Open(Stream stream)
    {
        XlsxStream.ValidateReadable(stream);
        return new ExcelWorkbookReader(stream);
    }

    /// <summary>Opens a seekable readable XLSX stream with explicit package and per-sheet resource limits.</summary>
    public static ExcelWorkbookReader Open(Stream stream, ExcelReadOptions options)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateReadOptions(options);
        return new ExcelWorkbookReader(stream, options);
    }

    /// <summary>Writes public scalar properties from each item to the first worksheet of an XLSX file.</summary>
    public static void Write<T>(string path, IEnumerable<T> rows)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(rows);
#else
        if (rows is null)
        {
            throw new ArgumentNullException(nameof(rows));
        }
#endif

        new XlsxWriter<T>(path).Write(rows);
    }

    /// <summary>Writes public scalar properties to a seekable writable stream, truncating it first and leaving it open at its end.</summary>
    public static void Write<T>(Stream stream, IEnumerable<T> rows)
    {
        ValidateWriteArguments(stream, rows);
        new XlsxWriter<T>(stream).Write(rows);
    }

    /// <summary>Writes rows to a stream with cooperative cancellation during row iteration.</summary>
    public static void Write<T>(Stream stream, IEnumerable<T> rows, CancellationToken cancellationToken)
    {
        ValidateWriteArguments(stream, rows);
        new XlsxWriter<T>(stream).Write(rows, cancellationToken: cancellationToken);
    }

    /// <summary>Writes rows with presentation options for the generated workbook.</summary>
    public static void Write<T>(string path, IEnumerable<T> rows, Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateWriteArguments(path, rows);
        new XlsxWriter<T>(path).Write(rows, options: ConfigureOptions(configure));
    }

    /// <summary>Writes rows with presentation options to a seekable writable stream without taking ownership.</summary>
    public static void Write<T>(Stream stream, IEnumerable<T> rows, Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateWriteArguments(stream, rows);
        new XlsxWriter<T>(stream).Write(rows, options: ConfigureOptions(configure));
    }

    /// <summary>Writes rows using an explicit typed schema.</summary>
    public static void Write<T>(string path, IEnumerable<T> rows, ExcelSchema<T> schema)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(rows);
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (rows is null)
        {
            throw new ArgumentNullException(nameof(rows));
        }

        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif

        new XlsxWriter<T>(path).Write(rows, schema);
    }

    /// <summary>Writes schema-mapped rows to a seekable writable stream without taking ownership.</summary>
    public static void Write<T>(Stream stream, IEnumerable<T> rows, ExcelSchema<T> schema)
    {
        ValidateWriteArguments(stream, rows);
        ValidateSchema(schema);
        new XlsxWriter<T>(stream).Write(rows, schema);
    }

    /// <summary>Writes schema-mapped rows to a stream with cooperative cancellation during row iteration.</summary>
    public static void Write<T>(Stream stream, IEnumerable<T> rows, ExcelSchema<T> schema, CancellationToken cancellationToken)
    {
        ValidateWriteArguments(stream, rows);
        ValidateSchema(schema);
        new XlsxWriter<T>(stream).Write(rows, schema, cancellationToken: cancellationToken);
    }

    /// <summary>Writes schema-mapped rows with presentation options for the generated workbook.</summary>
    public static void Write<T>(
        string path,
        IEnumerable<T> rows,
        ExcelSchema<T> schema,
        Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateWriteArguments(path, rows);

#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif

        new XlsxWriter<T>(path).Write(rows, schema, ConfigureOptions(configure));
    }

    /// <summary>Writes schema-mapped rows with presentation options to a seekable writable stream.</summary>
    public static void Write<T>(
        Stream stream,
        IEnumerable<T> rows,
        ExcelSchema<T> schema,
        Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateWriteArguments(stream, rows);
        ValidateSchema(schema);
        new XlsxWriter<T>(stream).Write(rows, schema, ConfigureOptions(configure));
    }

    /// <summary>Writes schema-mapped rows with an immutable per-operation schema overlay.</summary>
    public static void Write<T>(string path, IEnumerable<T> rows, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        ValidateWriteArguments(path, rows);
        ValidateSchemaAndOverlay(schema, overlay);
        new XlsxWriter<T>(path).Write(rows, schema, overlay: overlay);
    }

    /// <summary>Writes schema-mapped rows with a runtime overlay to a seekable writable stream.</summary>
    public static void Write<T>(Stream stream, IEnumerable<T> rows, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        ValidateWriteArguments(stream, rows);
        ValidateSchemaAndOverlay(schema, overlay);
        new XlsxWriter<T>(stream).Write(rows, schema, overlay: overlay);
    }

    /// <summary>Writes schema-mapped rows with a runtime overlay to a stream with cooperative cancellation.</summary>
    public static void Write<T>(Stream stream, IEnumerable<T> rows, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay, CancellationToken cancellationToken)
    {
        ValidateWriteArguments(stream, rows);
        ValidateSchemaAndOverlay(schema, overlay);
        new XlsxWriter<T>(stream).Write(rows, schema, overlay: overlay, cancellationToken: cancellationToken);
    }

    /// <summary>Writes schema-mapped rows with runtime schema and presentation configuration.</summary>
    public static void Write<T>(
        string path,
        IEnumerable<T> rows,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateWriteArguments(path, rows);
        ValidateSchemaAndOverlay(schema, overlay);
        new XlsxWriter<T>(path).Write(rows, schema, ConfigureOptions(configure), overlay);
    }

    /// <summary>Writes schema-mapped rows with runtime schema and presentation configuration to a seekable writable stream.</summary>
    public static void Write<T>(
        Stream stream,
        IEnumerable<T> rows,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateWriteArguments(stream, rows);
        ValidateSchemaAndOverlay(schema, overlay);
        new XlsxWriter<T>(stream).Write(rows, schema, ConfigureOptions(configure), overlay);
    }

    /// <summary>Creates an empty XLSX template from a typed schema.</summary>
    public static void CreateTemplate<T>(string path, ExcelSchema<T> schema)
    {
        ValidateTemplateArguments(path, schema);
        new XlsxTemplateWriter<T>(path).Write(schema);
    }

    /// <summary>Creates a template in a seekable writable stream, truncating it first and leaving it open at its end.</summary>
    public static void CreateTemplate<T>(Stream stream, ExcelSchema<T> schema)
    {
        ValidateTemplateArguments(stream, schema);
        new XlsxTemplateWriter<T>(stream).Write(schema);
    }

    /// <summary>Creates a template in a stream with cooperative cancellation.</summary>
    public static void CreateTemplate<T>(Stream stream, ExcelSchema<T> schema, CancellationToken cancellationToken)
    {
        ValidateTemplateArguments(stream, schema);
        new XlsxTemplateWriter<T>(stream).Write(schema, cancellationToken: cancellationToken);
    }

    /// <summary>Creates an empty XLSX template from a typed schema with presentation options.</summary>
    public static void CreateTemplate<T>(
        string path,
        ExcelSchema<T> schema,
        Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateTemplateArguments(path, schema);
        new XlsxTemplateWriter<T>(path).Write(schema, ConfigureOptions(configure));
    }

    /// <summary>Creates a configured template in a seekable writable stream without taking ownership.</summary>
    public static void CreateTemplate<T>(Stream stream, ExcelSchema<T> schema, Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateTemplateArguments(stream, schema);
        new XlsxTemplateWriter<T>(stream).Write(schema, ConfigureOptions(configure));
    }

    /// <summary>Creates an empty XLSX template with an immutable per-operation schema overlay.</summary>
    public static void CreateTemplate<T>(string path, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        ValidateTemplateArguments(path, schema);
        ValidateOverlay(overlay);
        new XlsxTemplateWriter<T>(path).Write(schema, overlay: overlay);
    }

    /// <summary>Creates a template with a runtime overlay in a seekable writable stream.</summary>
    public static void CreateTemplate<T>(Stream stream, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        ValidateTemplateArguments(stream, schema);
        ValidateOverlay(overlay);
        new XlsxTemplateWriter<T>(stream).Write(schema, overlay: overlay);
    }

    /// <summary>Creates a template with a runtime overlay in a stream with cooperative cancellation.</summary>
    public static void CreateTemplate<T>(Stream stream, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay, CancellationToken cancellationToken)
    {
        ValidateTemplateArguments(stream, schema);
        ValidateOverlay(overlay);
        new XlsxTemplateWriter<T>(stream).Write(schema, overlay: overlay, cancellationToken: cancellationToken);
    }

    /// <summary>Creates an empty XLSX template with runtime schema and presentation configuration.</summary>
    public static void CreateTemplate<T>(
        string path,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateTemplateArguments(path, schema);
        ValidateOverlay(overlay);
        new XlsxTemplateWriter<T>(path).Write(schema, ConfigureOptions(configure), overlay);
    }

    /// <summary>Creates a configured template with a runtime overlay in a seekable writable stream.</summary>
    public static void CreateTemplate<T>(Stream stream, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay, Action<ExcelWriteOptionsBuilder> configure)
    {
        ValidateTemplateArguments(stream, schema);
        ValidateOverlay(overlay);
        new XlsxTemplateWriter<T>(stream).Write(schema, ConfigureOptions(configure), overlay);
    }

    /// <summary>Reads the first worksheet of an XLSX file into typed rows.</summary>
    public static ExcelReadResult<T> Read<T>(string path)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        return new XlsxReader<T>(path).Read();
    }

    /// <summary>Reads the first worksheet from a seekable readable stream without taking ownership. Reading begins at position zero.</summary>
    public static ExcelReadResult<T> Read<T>(Stream stream)
    {
        XlsxStream.ValidateReadable(stream);
        return new XlsxReader<T>(stream).Read();
    }

    /// <summary>Reads the first worksheet with cooperative cancellation during row iteration.</summary>
    public static ExcelReadResult<T> Read<T>(Stream stream, CancellationToken cancellationToken)
    {
        XlsxStream.ValidateReadable(stream);
        return new XlsxReader<T>(stream).Read(cancellationToken: cancellationToken);
    }

    /// <summary>Reads the first worksheet using explicit per-operation read options.</summary>
    public static ExcelReadResult<T> Read<T>(string path, ExcelReadOptions options)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        ValidateReadOptions(options);
        return new XlsxReader<T>(path).Read(options: options);
    }

    /// <summary>Reads the first worksheet from a seekable readable stream using explicit read options.</summary>
    public static ExcelReadResult<T> Read<T>(Stream stream, ExcelReadOptions options)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateReadOptions(options);
        return new XlsxReader<T>(stream).Read(options: options);
    }

    /// <summary>Reads the first worksheet of an XLSX file using an explicit typed schema.</summary>
    public static ExcelReadResult<T> Read<T>(string path, ExcelSchema<T> schema)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif

        return new XlsxReader<T>(path).Read(schema);
    }

    /// <summary>Reads a schema-selected worksheet from a seekable readable stream without taking ownership.</summary>
    public static ExcelReadResult<T> Read<T>(Stream stream, ExcelSchema<T> schema)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateSchema(schema);
        return new XlsxReader<T>(stream).Read(schema);
    }

    /// <summary>Reads a schema-selected worksheet with cooperative cancellation during row iteration.</summary>
    public static ExcelReadResult<T> Read<T>(Stream stream, ExcelSchema<T> schema, CancellationToken cancellationToken)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateSchema(schema);
        return new XlsxReader<T>(stream).Read(schema, cancellationToken: cancellationToken);
    }

    /// <summary>Reads a worksheet using an explicit typed schema and per-operation read options.</summary>
    public static ExcelReadResult<T> Read<T>(string path, ExcelSchema<T> schema, ExcelReadOptions options)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        ValidateSchemaAndReadOptions(schema, options);
        return new XlsxReader<T>(path).Read(schema, options: options);
    }

    /// <summary>Reads a schema-selected worksheet from a seekable readable stream using explicit read options.</summary>
    public static ExcelReadResult<T> Read<T>(Stream stream, ExcelSchema<T> schema, ExcelReadOptions options)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateSchemaAndReadOptions(schema, options);
        return new XlsxReader<T>(stream).Read(schema, options: options);
    }

    /// <summary>Reads a worksheet using an explicit schema and immutable per-operation schema overlay.</summary>
    public static ExcelReadResult<T> Read<T>(string path, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        ValidateSchemaAndOverlay(schema, overlay);
        _ = RuntimeSchema<T>.Create(schema, overlay);
        return new XlsxReader<T>(path).Read(schema, overlay);
    }

    /// <summary>Reads a schema-selected worksheet with a runtime overlay from a seekable readable stream.</summary>
    public static ExcelReadResult<T> Read<T>(Stream stream, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateSchemaAndOverlay(schema, overlay);
        _ = RuntimeSchema<T>.Create(schema, overlay);
        return new XlsxReader<T>(stream).Read(schema, overlay);
    }

    /// <summary>Reads a worksheet using an explicit schema, runtime overlay, and read options.</summary>
    public static ExcelReadResult<T> Read<T>(
        string path,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        ExcelReadOptions options)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        ValidateSchemaAndOverlay(schema, overlay);
        ValidateReadOptions(options);
        _ = RuntimeSchema<T>.Create(schema, overlay);
        return new XlsxReader<T>(path).Read(schema, overlay, options);
    }

    /// <summary>Reads a schema-selected worksheet with a runtime overlay and read options from a seekable readable stream.</summary>
    public static ExcelReadResult<T> Read<T>(
        Stream stream,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        ExcelReadOptions options)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateSchemaAndOverlay(schema, overlay);
        ValidateReadOptions(options);
        _ = RuntimeSchema<T>.Create(schema, overlay);
        return new XlsxReader<T>(stream).Read(schema, overlay, options);
    }

    /// <summary>Reads a schema-selected worksheet with overlay, options, and cooperative cancellation.</summary>
    public static ExcelReadResult<T> Read<T>(
        Stream stream,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        ExcelReadOptions options,
        CancellationToken cancellationToken)
    {
        XlsxStream.ValidateReadable(stream);
        ValidateSchemaAndOverlay(schema, overlay);
        ValidateReadOptions(options);
        _ = RuntimeSchema<T>.Create(schema, overlay);
        return new XlsxReader<T>(stream).Read(schema, overlay, options, cancellationToken);
    }

    private static ExcelSchema<T> BuildAttributeSchema<T>(Action<ExcelSchemaBuilder<T>>? configure)
    {
        var type = typeof(T);
        var properties = type
            .GetProperties(BindingFlags.Instance | BindingFlags.Static | BindingFlags.Public)
            .Select(property => new AttributeProperty(
                property,
                property.GetCustomAttribute<ExcelColumnAttribute>(inherit: true),
                property.GetCustomAttribute<ExcelIgnoreAttribute>(inherit: true)))
            .ToArray();

        foreach (var property in properties.Where(property => property.Column is not null || property.Ignore is not null))
        {
            if (property.Column is not null && property.Ignore is not null)
            {
                throw AttributeConfigurationError<T>(property.Property, "cannot declare both ExcelColumnAttribute and ExcelIgnoreAttribute");
            }

            if (property.Property.GetMethod?.IsStatic == true || property.Property.SetMethod?.IsStatic == true ||
                property.Property.GetIndexParameters().Length != 0 || !property.Property.CanRead || property.Property.SetMethod?.IsPublic != true)
            {
                throw AttributeConfigurationError<T>(property.Property, "must be a public readable and writable instance property");
            }

            if (property.Column?.Header is not null && string.IsNullOrWhiteSpace(property.Column.Header))
            {
                throw AttributeConfigurationError<T>(property.Property, "has an empty or whitespace ExcelColumnAttribute header");
            }

            if (property.Column?.Format is not null && string.IsNullOrWhiteSpace(property.Column.Format))
            {
                throw AttributeConfigurationError<T>(property.Property, "has an empty or whitespace ExcelColumnAttribute format");
            }
        }

        var builder = new ExcelSchemaBuilder<T>();
        foreach (var property in properties
                     .Where(property => property.Property.GetMethod?.IsStatic != true && property.Property.SetMethod?.IsStatic != true)
                     .Where(property => property.Property.CanRead && property.Property.SetMethod?.IsPublic == true && property.Property.GetIndexParameters().Length == 0)
                     .OrderBy(property => property.Column?.HasOrder == true ? 0 : 1)
                     .ThenBy(property => property.Column?.HasOrder == true ? property.Column.Order : int.MaxValue)
                     .ThenBy(property => property.Property.MetadataToken))
        {
            builder.AddAttributeColumn(property.Property, property.Column, property.Ignore is not null);
        }

        configure?.Invoke(builder);
        return builder.Build();
    }

    private static ArgumentException AttributeConfigurationError<T>(PropertyInfo property, string reason) => new(
        $"Property '{property.Name}' on type '{typeof(T).FullName}' {reason}.");

    private sealed class AttributeProperty
    {
        internal AttributeProperty(PropertyInfo property, ExcelColumnAttribute? column, ExcelIgnoreAttribute? ignore)
        {
            Property = property;
            Column = column;
            Ignore = ignore;
        }

        internal PropertyInfo Property { get; }

        internal ExcelColumnAttribute? Column { get; }

        internal ExcelIgnoreAttribute? Ignore { get; }
    }

    private static ExcelWriteOptions ConfigureOptions(Action<ExcelWriteOptionsBuilder> configure)
    {
        if (configure is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(configure);
#else
            throw new ArgumentNullException(nameof(configure));
#endif
        }

        var builder = new ExcelWriteOptionsBuilder();
        configure(builder);
        return builder.Build();
    }

    private static void ValidateWriteArguments<T>(string path, IEnumerable<T> rows)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(rows);
#else
        if (rows is null)
        {
            throw new ArgumentNullException(nameof(rows));
        }
#endif
    }

    private static void ValidateWriteArguments<T>(Stream stream, IEnumerable<T> rows)
    {
        XlsxStream.ValidateWritable(stream);
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(rows);
#else
        if (rows is null)
        {
            throw new ArgumentNullException(nameof(rows));
        }
#endif
    }

    private static void ValidateTemplateArguments<T>(string path, ExcelSchema<T> schema)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif
    }

    private static void ValidateTemplateArguments<T>(Stream stream, ExcelSchema<T> schema)
    {
        XlsxStream.ValidateWritable(stream);
        ValidateSchema(schema);
    }

    private static void ValidateSchema<T>(ExcelSchema<T> schema)
    {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif
    }

    private static void ValidateSchemaAndOverlay<T>(ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        ValidateSchema(schema);
        ValidateOverlay(overlay);
    }

    private static void ValidateOverlay<T>(ExcelSchemaOverlay<T> overlay)
    {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(overlay);
#else
        if (overlay is null)
        {
            throw new ArgumentNullException(nameof(overlay));
        }
#endif
    }

    private static void ValidateSchemaAndReadOptions<T>(ExcelSchema<T> schema, ExcelReadOptions options)
    {
        ValidateSchema(schema);
        ValidateReadOptions(options);
    }

    private static void ValidateReadOptions(ExcelReadOptions options)
    {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(options);
#else
        if (options is null)
        {
            throw new ArgumentNullException(nameof(options));
        }
#endif
    }
}
