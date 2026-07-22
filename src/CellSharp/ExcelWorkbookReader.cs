using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using CellSharp.Internal;

namespace CellSharp;

/// <summary>Reads multiple typed worksheets from one open XLSX workbook.</summary>
public sealed class ExcelWorkbookReader : IDisposable
{
    private readonly SpreadsheetDocument _document;
    private readonly string _source;
    private readonly ExcelReadOptions _options;
    private bool _disposed;

    internal ExcelWorkbookReader(string path, ExcelReadOptions? options = null)
    {
        _options = options ?? ExcelReadOptions.Default;
        _source = Path.GetFullPath(path);
        if (!File.Exists(_source))
        {
            throw new FileNotFoundException($"The file '{_source}' does not exist.", _source);
        }

        _document = XlsxStream.Open(_source, _options);
    }

    internal ExcelWorkbookReader(Stream stream, ExcelReadOptions? options = null)
    {
        _options = options ?? ExcelReadOptions.Default;
        XlsxStream.ValidateReadable(stream);
        _source = "the supplied stream";
        _document = XlsxStream.Open(stream, _options);
    }

    /// <summary>Gets the names of readable worksheets in their workbook order.</summary>
    /// <remarks>Internal CellSharp support worksheets are excluded.</remarks>
    public IReadOnlyList<string> WorksheetNames
    {
        get
        {
            ThrowIfDisposed();
            return PublicSheets()
                .Select(sheet => sheet.Name?.Value ?? string.Empty)
                .ToArray();
        }
    }

    /// <summary>Reads the worksheet selected by the schema's SheetName.</summary>
    public ExcelReadResult<T> Read<T>(ExcelSchema<T> schema)
    {
        ThrowIfDisposed();
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif

        return XlsxReader<T>.Read(WorkbookPart(), SheetByName(schema.SheetName), schema, options: _options);
    }

    /// <summary>Reads the worksheet selected by the schema using a per-operation schema overlay.</summary>
    public ExcelReadResult<T> Read<T>(ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        ThrowIfDisposed();
        Validate(schema, overlay);
        return XlsxReader<T>.Read(WorkbookPart(), SheetByName(schema.SheetName), schema, overlay, _options);
    }

    /// <summary>Reads the worksheet selected by the schema using per-operation read options.</summary>
    public ExcelReadResult<T> Read<T>(ExcelSchema<T> schema, ExcelReadOptions options)
    {
        ThrowIfDisposed();
        Validate(schema, options);
        return XlsxReader<T>.Read(WorkbookPart(), SheetByName(schema.SheetName), schema, options: options);
    }

    /// <summary>Reads the worksheet selected by the schema with cooperative cancellation during row iteration.</summary>
    public ExcelReadResult<T> Read<T>(ExcelSchema<T> schema, CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif
        return XlsxReader<T>.Read(WorkbookPart(), SheetByName(schema.SheetName), schema, options: _options, cancellationToken: cancellationToken);
    }

    /// <summary>Reads the worksheet selected by the schema using a runtime overlay and read options.</summary>
    public ExcelReadResult<T> Read<T>(ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay, ExcelReadOptions options)
    {
        ThrowIfDisposed();
        Validate(schema, overlay);
        Validate(options);
        return XlsxReader<T>.Read(WorkbookPart(), SheetByName(schema.SheetName), schema, overlay, options);
    }

    /// <summary>Reads a zero-based public worksheet index after verifying it matches the schema's SheetName.</summary>
    public ExcelReadResult<T> ReadAt<T>(int sheetIndex, ExcelSchema<T> schema)
    {
        ThrowIfDisposed();
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif

        var sheets = PublicSheets();
        if (sheetIndex < 0 || sheetIndex >= sheets.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Worksheet indexes are zero-based and must identify a public worksheet.");
        }

        var sheet = sheets[sheetIndex];
        if (!string.Equals(sheet.Name?.Value, schema.SheetName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Worksheet index {sheetIndex} identifies '{sheet.Name?.Value}', which does not match schema sheet '{schema.SheetName}'.");
        }

        return XlsxReader<T>.Read(WorkbookPart(), sheet, schema, options: _options);
    }

    /// <summary>Reads a public worksheet index using a per-operation schema overlay.</summary>
    public ExcelReadResult<T> ReadAt<T>(int sheetIndex, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay)
    {
        ThrowIfDisposed();
        Validate(schema, overlay);
        var sheets = PublicSheets();
        if (sheetIndex < 0 || sheetIndex >= sheets.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Worksheet indexes are zero-based and must identify a public worksheet.");
        }

        var sheet = sheets[sheetIndex];
        if (!string.Equals(sheet.Name?.Value, schema.SheetName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Worksheet index {sheetIndex} identifies '{sheet.Name?.Value}', which does not match schema sheet '{schema.SheetName}'.");
        }

        return XlsxReader<T>.Read(WorkbookPart(), sheet, schema, overlay, _options);
    }

    /// <summary>Reads a public worksheet index using per-operation read options.</summary>
    public ExcelReadResult<T> ReadAt<T>(int sheetIndex, ExcelSchema<T> schema, ExcelReadOptions options)
    {
        ThrowIfDisposed();
        Validate(schema, options);
        return ReadAtCore(sheetIndex, schema, options: options);
    }

    /// <summary>Reads a public worksheet index with cooperative cancellation during row iteration.</summary>
    public ExcelReadResult<T> ReadAt<T>(int sheetIndex, ExcelSchema<T> schema, CancellationToken cancellationToken)
    {
        ThrowIfDisposed();
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif
        return ReadAtCore(sheetIndex, schema, cancellationToken: cancellationToken);
    }

    /// <summary>Reads a public worksheet index using a runtime overlay and read options.</summary>
    public ExcelReadResult<T> ReadAt<T>(int sheetIndex, ExcelSchema<T> schema, ExcelSchemaOverlay<T> overlay, ExcelReadOptions options)
    {
        ThrowIfDisposed();
        Validate(schema, overlay);
        Validate(options);
        return ReadAtCore(sheetIndex, schema, overlay, options);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (_disposed)
        {
            return;
        }

        _document.Dispose();
        _disposed = true;
    }

    private WorkbookPart WorkbookPart() => _document.WorkbookPart
        ?? throw new InvalidOperationException("The XLSX file has no workbook part.");

    private Sheet SheetByName(string sheetName)
    {
        var sheet = Workbook().Sheets?.Elements<Sheet>().FirstOrDefault(candidate => string.Equals(
            candidate.Name?.Value,
            sheetName,
            StringComparison.OrdinalIgnoreCase));
        if (sheet is null)
        {
            throw new InvalidOperationException($"Worksheet '{sheetName}' was not found in '{_source}'.");
        }

        if (XlsxDataValidationWriter.IsInternalSheet(sheet))
        {
            throw new InvalidOperationException("CellSharp internal worksheets cannot be read as application data.");
        }

        return sheet;
    }

    private Sheet[] PublicSheets()
    {
        var sheets = Workbook().Sheets?.Elements<Sheet>()
            .Where(sheet => !XlsxDataValidationWriter.IsInternalSheet(sheet))
            .ToArray();
        if (sheets is null || sheets.Length == 0)
        {
            throw new InvalidOperationException("The XLSX file has no readable worksheet.");
        }

        return sheets;
    }

    private Workbook Workbook() => WorkbookPart().Workbook
        ?? throw new InvalidOperationException("The XLSX file has no workbook.");

    private void ThrowIfDisposed()
    {
#if NET8_0_OR_GREATER
        ObjectDisposedException.ThrowIf(_disposed, this);
#else
        if (_disposed)
        {
            throw new ObjectDisposedException(nameof(ExcelWorkbookReader));
        }
#endif
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

    private ExcelReadResult<T> ReadAtCore<T>(
        int sheetIndex,
        ExcelSchema<T> schema,
        ExcelSchemaOverlay<T>? overlay = null,
        ExcelReadOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        var sheets = PublicSheets();
        if (sheetIndex < 0 || sheetIndex >= sheets.Length)
        {
            throw new ArgumentOutOfRangeException(nameof(sheetIndex), "Worksheet indexes are zero-based and must identify a public worksheet.");
        }

        var sheet = sheets[sheetIndex];
        if (!string.Equals(sheet.Name?.Value, schema.SheetName, StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException(
                $"Worksheet index {sheetIndex} identifies '{sheet.Name?.Value}', which does not match schema sheet '{schema.SheetName}'.");
        }

        return XlsxReader<T>.Read(WorkbookPart(), sheet, schema, overlay, options ?? _options, cancellationToken);
    }

    private static void Validate<T>(ExcelSchema<T> schema, ExcelReadOptions options)
    {
#if NET8_0_OR_GREATER
        ArgumentNullException.ThrowIfNull(schema);
#else
        if (schema is null)
        {
            throw new ArgumentNullException(nameof(schema));
        }
#endif
        Validate(options);
    }

    private static void Validate(ExcelReadOptions options)
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
