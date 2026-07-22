using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal sealed class XlsxTemplateWriter<T>
{
    private readonly string? _path;
    private readonly Stream? _stream;

    internal XlsxTemplateWriter(string path)
    {
        _path = Path.GetFullPath(path);
    }

    internal XlsxTemplateWriter(Stream stream)
    {
        _stream = stream;
    }

    internal void Write(ExcelSchema<T> schema, ExcelWriteOptions? options = null, ExcelSchemaOverlay<T>? overlay = null, CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var writeOptions = options ?? ExcelWriteOptions.Default;
        _ = ExportProperty.For(schema, overlay);

        if (_stream is not null)
        {
            using (var document = XlsxStream.Create(_stream))
            {
                WriteDocument(document, schema, writeOptions, overlay, cancellationToken);
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
        WriteDocument(pathDocument, schema, writeOptions, overlay, cancellationToken);
    }

    private static void WriteDocument(
        SpreadsheetDocument document,
        ExcelSchema<T> schema,
        ExcelWriteOptions writeOptions,
        ExcelSchemaOverlay<T>? overlay,
        CancellationToken cancellationToken)
    {
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var styles = new WorkbookStyleCatalog(workbookPart);
        var worksheet = XlsxWorksheetWriter.WriteTemplate(workbookPart, styles, schema, writeOptions, overlay, cancellationToken);
        XlsxDataValidationWriter.Apply(workbookPart, [worksheet], cancellationToken);
        XlsxWorksheetSettingsWriter.Apply(workbookPart, [worksheet], cancellationToken);
        XlsxTableWriter.Apply([worksheet], cancellationToken);
        styles.Save();
        worksheet.WorksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }
}
