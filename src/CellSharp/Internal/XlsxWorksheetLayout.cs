using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxWorksheetLayout
{
    internal static void AddFrozenHeaderView(Worksheet worksheet, ResolvedWorksheetSettings settings)
    {
        if (settings.FreezeRows == 0 && settings.FreezeColumns == 0)
        {
            return;
        }

        var topLeftCell = $"{ColumnName(settings.FreezeColumns + 1)}{settings.FreezeRows + 1}";
        var activePane = settings.FreezeRows > 0 && settings.FreezeColumns > 0
            ? PaneValues.BottomRight
            : settings.FreezeRows > 0
                ? PaneValues.BottomLeft
                : PaneValues.TopRight;

        worksheet.AppendChild(new SheetViews(
            new SheetView(
                new Pane
                {
                    HorizontalSplit = settings.FreezeColumns == 0 ? null : (double?)settings.FreezeColumns,
                    VerticalSplit = settings.FreezeRows == 0 ? null : (double?)settings.FreezeRows,
                    TopLeftCell = topLeftCell,
                    ActivePane = activePane,
                    State = PaneStateValues.Frozen,
                })
            {
                WorkbookViewId = 0U,
            }));
    }

    internal static Row HeaderRow(IEnumerable<ExportProperty> properties, WorkbookStyleCatalog styles, ExcelWriteOptions options)
    {
        var row = new Row();
        foreach (var property in properties)
        {
            row.AppendChild(CellValueWriter.Create(property.Header, styles.HeaderStyleIndex(property, options)));
        }

        return row;
    }

    internal static Columns ConfiguredColumns(
        IReadOnlyList<ExportProperty> properties,
        IReadOnlyList<int> widths,
        bool autoFitColumns)
    {
        var columns = new Columns();
        for (var index = 0; index < widths.Count; index++)
        {
            if (properties[index].Width is null && !autoFitColumns)
            {
                continue;
            }

            columns.AppendChild(Column(properties[index], index, widths[index], autoFitColumns));
        }

        return columns;
    }

    internal static Columns TemplateColumns(
        IReadOnlyList<ExportProperty> properties,
        IReadOnlyList<int> widths,
        bool autoFitColumns,
        WorkbookStyleCatalog styles,
        ExcelWriteOptions options)
    {
        var columns = new Columns();
        for (var index = 0; index < properties.Count; index++)
        {
            var column = Column(properties[index], index, widths[index], autoFitColumns);
            column.Style = styles.TemplateDataStyleIndex(properties[index], options);
            columns.AppendChild(column);
        }

        return columns;
    }

    internal static void AddSheet(WorkbookPart workbookPart, WorksheetPart worksheetPart, string sheetName)
    {
        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("The workbook part has no workbook.");
        var sheets = workbook.Sheets ?? workbook.AppendChild(new Sheets());
        var sheetId = sheets.Elements<Sheet>().Select(sheet => sheet.SheetId?.Value ?? 0U).DefaultIfEmpty(0U).Max() + 1U;
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = sheetName,
        });
    }

    private static Column Column(ExportProperty property, int index, int estimatedWidth, bool autoFitColumns)
    {
        var column = new Column
        {
            Min = (uint)(index + 1),
            Max = (uint)(index + 1),
        };

        if (property.Width is not null || autoFitColumns)
        {
            column.Width = property.Width ?? Math.Max(10D, Math.Min(60D, estimatedWidth + 2D));
            column.CustomWidth = true;
        }

        return column;
    }

    internal static string ColumnName(int columnNumber)
    {
        var name = string.Empty;
        var value = columnNumber;
        while (value > 0)
        {
            value--;
            name = (char)('A' + (value % 26)) + name;
            value /= 26;
        }

        return name;
    }
}
