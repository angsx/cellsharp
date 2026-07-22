using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;

namespace CellSharp.Internal;

internal interface IWorkbookSheetDefinition
{
    string SheetName { get; }

    bool HasRows { get; }

    void Validate();

    WorksheetValidationContext Write(WorkbookPart workbookPart, WorkbookStyleCatalog styles, CancellationToken cancellationToken);

    WorksheetValidationContext CreateTemplate(WorkbookPart workbookPart, WorkbookStyleCatalog styles, CancellationToken cancellationToken);
}

internal sealed class WorkbookSheetDefinition<T> : IWorkbookSheetDefinition
{
    private readonly IEnumerable<T>? _rows;
    private readonly ExcelSchema<T> _schema;
    private readonly ExcelWriteOptions _options;
    private readonly ExcelSchemaOverlay<T>? _overlay;
    private readonly WorksheetLayoutDefinition? _layout;

    internal WorkbookSheetDefinition(
        IEnumerable<T> rows,
        ExcelSchema<T> schema,
        ExcelWriteOptions options,
        ExcelSchemaOverlay<T>? overlay = null,
        WorksheetLayoutDefinition? layout = null)
    {
        _rows = rows;
        _schema = schema;
        _options = options;
        _overlay = overlay;
        _layout = layout;
    }

    internal WorkbookSheetDefinition(ExcelSchema<T> schema, ExcelWriteOptions options, ExcelSchemaOverlay<T>? overlay = null, WorksheetLayoutDefinition? layout = null)
    {
        _schema = schema;
        _options = options;
        _overlay = overlay;
        _layout = layout;
    }

    public string SheetName => _schema.SheetName;

    public bool HasRows => _rows is not null;

    public void Validate() => _ = ExportProperty.For(_schema, _overlay);

    public WorksheetValidationContext Write(WorkbookPart workbookPart, WorkbookStyleCatalog styles, CancellationToken cancellationToken) => _rows is null
        ? throw new InvalidOperationException("The worksheet does not have export rows.")
        : XlsxWorksheetWriter.WriteData(workbookPart, styles, _rows, _schema, _options, _overlay, _layout, cancellationToken);

    public WorksheetValidationContext CreateTemplate(WorkbookPart workbookPart, WorkbookStyleCatalog styles, CancellationToken cancellationToken) => XlsxWorksheetWriter.WriteTemplate(
        workbookPart,
        styles,
        _schema,
        _options,
        _overlay,
        layout: _layout,
        cancellationToken: cancellationToken);
}

internal sealed class LayoutWorkbookSheetDefinition : IWorkbookSheetDefinition
{
    private readonly WorksheetLayoutDefinition _layout;
    internal LayoutWorkbookSheetDefinition(WorksheetLayoutDefinition layout) => _layout = layout;
    public string SheetName => _layout.Name;
    public bool HasRows => false;
    public void Validate() { }
    public WorksheetValidationContext Write(WorkbookPart workbookPart, WorkbookStyleCatalog styles, CancellationToken cancellationToken) => XlsxWorksheetWriter.WriteLayout(workbookPart, styles, _layout, cancellationToken);
    public WorksheetValidationContext CreateTemplate(WorkbookPart workbookPart, WorkbookStyleCatalog styles, CancellationToken cancellationToken) => XlsxWorksheetWriter.WriteLayout(workbookPart, styles, _layout, cancellationToken);
}

internal static class XlsxWorkbookWriter
{
    internal static void Write(string path, IReadOnlyList<IWorkbookSheetDefinition> sheets, bool createTemplate, CancellationToken cancellationToken = default)
    {
        var fullPath = Path.GetFullPath(path);
        var directory = Path.GetDirectoryName(fullPath);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException($"The directory for '{fullPath}' does not exist.");
        }

        foreach (var sheet in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheet.Validate();
        }

        using var document = SpreadsheetDocument.Create(fullPath, SpreadsheetDocumentType.Workbook);
        WriteDocument(document, sheets, createTemplate, cancellationToken);
    }

    internal static void Write(Stream stream, IReadOnlyList<IWorkbookSheetDefinition> sheets, bool createTemplate, CancellationToken cancellationToken = default)
    {
        foreach (var sheet in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            sheet.Validate();
        }

        using (var document = XlsxStream.Create(stream))
        {
            WriteDocument(document, sheets, createTemplate, cancellationToken);
        }

        XlsxStream.CompleteWrite(stream);
    }

    private static void WriteDocument(
        SpreadsheetDocument document,
        IReadOnlyList<IWorkbookSheetDefinition> sheets,
        bool createTemplate,
        CancellationToken cancellationToken)
    {
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new DocumentFormat.OpenXml.Spreadsheet.Workbook();
        var styles = new WorkbookStyleCatalog(workbookPart);
        var worksheets = new List<WorksheetValidationContext>(sheets.Count);
        foreach (var sheet in sheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            worksheets.Add(createTemplate
                ? sheet.CreateTemplate(workbookPart, styles, cancellationToken)
                : sheet.Write(workbookPart, styles, cancellationToken));
        }

        XlsxDataValidationWriter.Apply(workbookPart, worksheets, cancellationToken);
        XlsxWorksheetSettingsWriter.Apply(workbookPart, worksheets, cancellationToken);
        XlsxDefinedNamesWriter.Apply(workbookPart, worksheets);
        XlsxHeaderFooterWriter.Apply(worksheets);
        XlsxPageBreakWriter.Apply(worksheets);
        XlsxImageWriter.Apply(worksheets);
        XlsxCommentsWriter.Apply(worksheets);
        XlsxTableWriter.Apply(worksheets, cancellationToken);
        if (!createTemplate)
        {
            XlsxWorksheetWriter.RequestFormulaRecalculation(workbookPart, worksheets);
        }

        styles.Save();
        foreach (var worksheet in worksheets)
        {
            worksheet.WorksheetPart.Worksheet!.Save();
        }

        workbookPart.Workbook!.Save();
    }
}
