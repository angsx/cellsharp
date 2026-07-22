using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxWorksheetSettingsWriter
{
    internal static void Apply(WorkbookPart workbookPart, IReadOnlyList<WorksheetValidationContext> worksheets, CancellationToken cancellationToken = default)
    {
        foreach (var worksheet in worksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            ApplyWorksheetSettings(worksheet);
        }

        ApplyPrintTitles(workbookPart, worksheets, cancellationToken);
    }

    private static void ApplyWorksheetSettings(WorksheetValidationContext context)
    {
        var worksheet = context.Worksheet;
        if (context.Settings.AutoFilter && context.Table is null)
        {
            var filter = new AutoFilter
            {
                Reference = $"{XlsxWorksheetLayout.ColumnName(context.DataStartColumn)}{context.DataStartRow}:{XlsxWorksheetLayout.ColumnName(context.DataStartColumn + context.Properties.Count - 1)}{context.DataStartRow + context.DataRowCount}",
            };
            InsertAutoFilter(worksheet, filter);
        }

        if (context.Settings.HasFitToPage)
        {
            var properties = worksheet.GetFirstChild<SheetProperties>();
            if (properties is null)
            {
                properties = new SheetProperties();
                worksheet.InsertAt(properties, 0);
            }

            properties.PageSetupProperties = new PageSetupProperties { FitToPage = true };
        }

        if (context.Settings.PrintGridlines)
        {
            worksheet.AppendChild(new PrintOptions { GridLines = true });
        }

        if (context.Settings.Landscape is not null || context.Settings.HasFitToPage)
        {
            var pageSetup = new PageSetup
            {
                FitToWidth = context.Settings.FitToWidth is null ? null : (uint?)context.Settings.FitToWidth.Value,
                FitToHeight = context.Settings.FitToHeight is null ? null : (uint?)context.Settings.FitToHeight.Value,
            };
            if (context.Settings.Landscape is not null)
            {
                pageSetup.Orientation = context.Settings.Landscape.Value
                    ? OrientationValues.Landscape
                    : OrientationValues.Portrait;
            }

            worksheet.AppendChild(pageSetup);
        }
    }

    private static void InsertAutoFilter(Worksheet worksheet, AutoFilter filter)
    {
        // autoFilter must precede merges and every later worksheet feature. Layout is written
        // before worksheet settings, so inserting at the end would make combined workbooks invalid.
        var next = worksheet.ChildElements.FirstOrDefault(element => element is MergeCells
            or PhoneticProperties
            or ConditionalFormatting
            or DataValidations
            or Hyperlinks
            or PrintOptions
            or PageMargins
            or PageSetup
            or HeaderFooter
            or RowBreaks
            or ColumnBreaks
            or Drawing
            or LegacyDrawing
            or TableParts);
        if (next is null)
        {
            worksheet.AppendChild(filter);
        }
        else
        {
            worksheet.InsertBefore(filter, next);
        }
    }

    private static void ApplyPrintTitles(
        WorkbookPart workbookPart,
        IReadOnlyList<WorksheetValidationContext> worksheets,
        CancellationToken cancellationToken)
    {
        var titleWorksheets = worksheets.Where(worksheet => worksheet.Settings.RepeatHeaderRowOnPrint).ToArray();
        if (titleWorksheets.Length == 0)
        {
            return;
        }

        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("The workbook part has no workbook.");
        var sheets = workbook.Sheets!.Elements<Sheet>().ToArray();
        var definedNames = workbook.DefinedNames ?? workbook.AppendChild(new DefinedNames());
        foreach (var worksheet in titleWorksheets)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var relationshipId = workbookPart.GetIdOfPart(worksheet.WorksheetPart);
            var localSheetId = Array.FindIndex(sheets, sheet => sheet.Id!.Value == relationshipId);
            if (localSheetId < 0)
            {
                throw new InvalidOperationException($"Worksheet '{worksheet.SheetName}' is not attached to the workbook.");
            }

            var escapedName = worksheet.SheetName.Replace("'", "''");
            definedNames.AppendChild(new DefinedName
            {
                Name = "_xlnm.Print_Titles",
                LocalSheetId = (uint)localSheetId,
                Text = $"'{escapedName}'!$1:$1",
            });
        }
    }
}
