using CellSharp.Internal;

namespace CellSharp;

/// <summary>Builds a typed XLSX workbook containing multiple worksheets.</summary>
public sealed class ExcelWorkbookBuilder
{
    private readonly List<IWorkbookSheetDefinition> _sheets = new();

    /// <summary>Adds a typed data worksheet to a workbook export.</summary>
    public ExcelWorkbookBuilder AddSheet<T>(
        IEnumerable<T> rows,
        ExcelSchema<T> schema,
        Action<ExcelWriteOptionsBuilder>? configure = null)
    {
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

        Add(new WorkbookSheetDefinition<T>(rows, schema, ConfigureOptions(configure)));
        return this;
    }

    /// <summary>Adds a typed data worksheet and configures arbitrary worksheet layout.</summary>
    public ExcelWorkbookBuilder AddSheet<T>(IEnumerable<T> rows, ExcelSchema<T> schema, Action<ExcelWorksheetBuilder> layout)
    {
        if (rows is null) throw new ArgumentNullException(nameof(rows));
        if (schema is null) throw new ArgumentNullException(nameof(schema));
        if (layout is null) throw new ArgumentNullException(nameof(layout));
        var builder = new ExcelWorksheetBuilder(schema.SheetName);
        layout(builder);
        Add(new WorkbookSheetDefinition<T>(rows, schema, ExcelWriteOptions.Default, null, builder.Definition));
        return this;
    }

    /// <summary>Adds a layout-only worksheet for a report without a typed dataset.</summary>
    public ExcelWorkbookBuilder AddSheet(string name, Action<ExcelWorksheetBuilder> configure)
    {
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelWorksheetBuilder(name);
        configure(builder);
        Add(new LayoutWorkbookSheetDefinition(builder.Definition));
        return this;
    }

    /// <summary>Adds a typed data worksheet with a per-operation schema overlay.</summary>
    public ExcelWorkbookBuilder AddSheet<T>(
        IEnumerable<T> rows,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        Action<ExcelWriteOptionsBuilder>? configure = null)
    {
        Validate(rows, schema, overlay);
        Add(new WorkbookSheetDefinition<T>(rows, schema, ConfigureOptions(configure), overlay));
        return this;
    }

    /// <summary>Adds a typed header-only worksheet to a workbook template.</summary>
    public ExcelWorkbookBuilder AddTemplateSheet<T>(
        ExcelSchema<T> schema,
        Action<ExcelWriteOptionsBuilder>? configure = null)
    {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif

        Add(new WorkbookSheetDefinition<T>(schema, ConfigureOptions(configure)));
        return this;
    }

    /// <summary>Adds a typed template worksheet and configures arbitrary worksheet layout.</summary>
    public ExcelWorkbookBuilder AddTemplateSheet<T>(ExcelSchema<T> schema, Action<ExcelWorksheetBuilder> layout)
    {
        if (schema is null) throw new ArgumentNullException(nameof(schema));
        if (layout is null) throw new ArgumentNullException(nameof(layout));
        var builder = new ExcelWorksheetBuilder(schema.SheetName);
        layout(builder);
        Add(new WorkbookSheetDefinition<T>(schema, ExcelWriteOptions.Default, null, builder.Definition));
        return this;
    }

    /// <summary>Adds a typed header-only worksheet with a per-operation schema overlay.</summary>
    public ExcelWorkbookBuilder AddTemplateSheet<T>(
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T> overlay,
        Action<ExcelWriteOptionsBuilder>? configure = null)
    {
        Validate(schema, overlay);
        Add(new WorkbookSheetDefinition<T>(schema, ConfigureOptions(configure), overlay));
        return this;
    }

    /// <summary>Writes all configured data worksheets to a new XLSX workbook.</summary>
    public void Write(string path)
    {
        Validate(path, expectedDataSheets: true);
        XlsxWorkbookWriter.Write(path, _sheets, createTemplate: false);
    }

    /// <summary>Writes all configured data worksheets to a seekable writable stream, truncating it first and leaving it open at its end.</summary>
    public void Write(Stream stream)
    {
        Validate(stream, expectedDataSheets: true);
        XlsxWorkbookWriter.Write(stream, _sheets, createTemplate: false);
    }

    /// <summary>Writes all configured data worksheets with cooperative cancellation between sheets and rows.</summary>
    public void Write(Stream stream, CancellationToken cancellationToken)
    {
        Validate(stream, expectedDataSheets: true);
        XlsxWorkbookWriter.Write(stream, _sheets, createTemplate: false, cancellationToken: cancellationToken);
    }

    /// <summary>Creates a header-only XLSX workbook from all configured template worksheets.</summary>
    public void CreateTemplate(string path)
    {
        Validate(path, expectedDataSheets: false);
        XlsxWorkbookWriter.Write(path, _sheets, createTemplate: true);
    }

    /// <summary>Creates all configured template worksheets in a seekable writable stream without taking ownership.</summary>
    public void CreateTemplate(Stream stream)
    {
        Validate(stream, expectedDataSheets: false);
        XlsxWorkbookWriter.Write(stream, _sheets, createTemplate: true);
    }

    /// <summary>Creates all configured template worksheets with cooperative cancellation.</summary>
    public void CreateTemplate(Stream stream, CancellationToken cancellationToken)
    {
        Validate(stream, expectedDataSheets: false);
        XlsxWorkbookWriter.Write(stream, _sheets, createTemplate: true, cancellationToken: cancellationToken);
    }

    private void Add(IWorkbookSheetDefinition sheet)
    {
        if (_sheets.Any(existing => string.Equals(existing.SheetName, sheet.SheetName, StringComparison.OrdinalIgnoreCase)))
        {
            throw new ArgumentException($"Worksheet name '{sheet.SheetName}' is already configured.", nameof(sheet));
        }

        _sheets.Add(sheet);
    }

    private void Validate(string path, bool expectedDataSheets)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("A file path is required.", nameof(path));
        }

        if (_sheets.Count == 0)
        {
            throw new InvalidOperationException("A workbook must include at least one worksheet.");
        }

        if (!expectedDataSheets && _sheets.Any(sheet => sheet.HasRows))
        {
            throw new InvalidOperationException("CreateTemplate requires template worksheets added with AddTemplateSheet.");
        }
    }

    private void Validate(Stream stream, bool expectedDataSheets)
    {
        XlsxStream.ValidateWritable(stream);
        ValidateSheets(expectedDataSheets);
    }

    private void ValidateSheets(bool expectedDataSheets)
    {
        if (_sheets.Count == 0)
        {
            throw new InvalidOperationException("A workbook must include at least one worksheet.");
        }

        if (!expectedDataSheets && _sheets.Any(sheet => sheet.HasRows))
        {
            throw new InvalidOperationException("CreateTemplate requires template worksheets added with AddTemplateSheet.");
        }
    }

    private static ExcelWriteOptions ConfigureOptions(Action<ExcelWriteOptionsBuilder>? configure)
    {
        if (configure is null)
        {
            return ExcelWriteOptions.Default;
        }

        var builder = new ExcelWriteOptionsBuilder();
        configure(builder);
        return builder.Build();
    }

    private static void Validate<T>(IEnumerable<T> rows, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(rows);
#else
        if (rows is null)
        {
            throw new ArgumentNullException(nameof(rows));
        }
#endif
        Validate(schema, overlay);
    }

    private static void Validate<T>(ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
        ArgumentNullException.ThrowIfNull(overlay);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }

        if (overlay is null)
        {
            throw new ArgumentNullException(nameof(overlay));
        }
#endif
    }
}
