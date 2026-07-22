using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal sealed class XlsxWriter<T>
{
    private readonly string? _path;
    private readonly Stream? _stream;

    internal XlsxWriter(string path)
    {
        _path = Path.GetFullPath(path);
    }

    internal XlsxWriter(Stream stream)
    {
        _stream = stream;
    }

    internal void Write(
        IEnumerable<T> rows,
        ExcelSchema<T>? schema = null,
        ExcelWriteOptions? options = null,
        ExcelSchemaOverlay<T>? overlay = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        _ = ExportProperty.For(schema, overlay);
        var writeOptions = options ?? ExcelWriteOptions.Default;

        if (_stream is not null)
        {
            using (var document = XlsxStream.Create(_stream))
            {
                WriteDocument(document, rows, schema, writeOptions, overlay, cancellationToken);
            }

            XlsxStream.CompleteWrite(_stream);
            return;
        }

        var directory = Path.GetDirectoryName(_path);
        if (string.IsNullOrEmpty(directory) || !Directory.Exists(directory))
        {
            throw new DirectoryNotFoundException($"The directory for '{_path}' does not exist.");
        }

        using var pathDocument = SpreadsheetDocument.Create(_path!, SpreadsheetDocumentType.Workbook);
        WriteDocument(pathDocument, rows, schema, writeOptions, overlay, cancellationToken);
    }

    private static void WriteDocument(
        SpreadsheetDocument document,
        IEnumerable<T> rows,
        ExcelSchema<T>? schema,
        ExcelWriteOptions writeOptions,
        ExcelSchemaOverlay<T>? overlay,
        CancellationToken cancellationToken)
    {
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var styles = new WorkbookStyleCatalog(workbookPart);
        var worksheet = XlsxWorksheetWriter.WriteData(workbookPart, styles, rows, schema, writeOptions, overlay, cancellationToken: cancellationToken);
        XlsxDataValidationWriter.Apply(workbookPart, [worksheet], cancellationToken);
        XlsxWorksheetSettingsWriter.Apply(workbookPart, [worksheet], cancellationToken);
        XlsxTableWriter.Apply([worksheet], cancellationToken);
        XlsxWorksheetWriter.RequestFormulaRecalculation(workbookPart, [worksheet]);

        styles.Save();
        worksheet.WorksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }
}
