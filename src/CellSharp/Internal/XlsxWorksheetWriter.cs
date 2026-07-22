using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxWorksheetWriter
{
    private const int MaximumDataRows = 1048575;
    private const int MaximumColumns = 16384;

    internal static void RequestFormulaRecalculation(WorkbookPart workbookPart, IEnumerable<WorksheetValidationContext> worksheets)
    {
        if (!worksheets.Any(worksheet => worksheet.HasFormulaCells))
        {
            return;
        }

        var workbook = workbookPart.Workbook
            ?? throw new InvalidOperationException("The workbook part has no workbook.");
        workbook.CalculationProperties = new CalculationProperties
        {
            FullCalculationOnLoad = true,
            ForceFullCalculation = true,
            CalculationOnSave = true,
        };
    }

    internal static WorksheetValidationContext WriteData<T>(
        WorkbookPart workbookPart,
        WorkbookStyleCatalog styles,
        IEnumerable<T> rows,
        ExcelSchema<T>? schema,
        ExcelWriteOptions options,
        ExcelSchemaOverlay<T>? overlay = null,
        WorksheetLayoutDefinition? layout = null,
        CancellationToken cancellationToken = default)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var properties = ExportProperty.For(schema, overlay);
        ValidateWorksheetShape(properties);
        var settings = schema?.WorksheetSettings.Resolve(options) ?? WorksheetSettingsDefinition.From(options);
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        var worksheet = new Worksheet();
        XlsxWorksheetLayout.AddFrozenHeaderView(worksheet, settings);
        worksheet.AppendChild(sheetData);
        worksheetPart.Worksheet = worksheet;
        var columnWidths = properties.Select(property => CellValueWriter.DisplayWidth(property.Header)).ToArray();

        var dataStart = layout?.DataStart;
        var startRow = dataStart?.FromRow ?? 1;
        var startColumn = dataStart?.FromColumn ?? 1;
        sheetData.AppendChild(HeaderRow(properties, styles.HeaderStyleIndex(options), startRow, startColumn));
        var dataRowIndex = 0;
        var hasFormulaCells = false;
        foreach (var item in rows)
        {
            cancellationToken.ThrowIfCancellationRequested();
            if (dataRowIndex >= MaximumDataRows)
            {
                throw new InvalidOperationException($"An XLSX worksheet cannot contain more than {MaximumDataRows} data rows plus its header.");
            }

            var row = new Row { RowIndex = (uint)(startRow + dataRowIndex + 1) };
            for (var columnIndex = 0; columnIndex < properties.Count; columnIndex++)
            {
                var property = properties[columnIndex];
                var formula = property.Formula is null
                    ? null
                    : Formula(property.Formula(new ExcelFormulaContext((uint)(startRow + dataRowIndex + 1), startColumn + columnIndex, schema?.SheetName ?? ExcelSheetName.Default)));
                hasFormulaCells |= formula is not null;
                // A formula is still the cell's source of truth, but its cached result makes a
                // workbook immediately readable before a spreadsheet application recalculates it.
                var value = property.ConvertValue(property.GetValue(item));
                columnWidths[columnIndex] = Math.Max(columnWidths[columnIndex], CellValueWriter.DisplayWidth(formula ?? value));
                var styleIndex = styles.DataStyleIndex(property, value, dataRowIndex, options);
                row.AppendChild(formula is null
                    ? CellValueWriter.Create(value, styleIndex)
                    : CellValueWriter.CreateFormula(formula, value, styleIndex));
            }

            var cells = row.Elements<Cell>().ToArray();
            for (var columnIndex = 0; columnIndex < cells.Length; columnIndex++) cells[columnIndex].CellReference = ExcelRangeReference.ColumnName(startColumn + columnIndex) + (startRow + dataRowIndex + 1);
            sheetData.AppendChild(row);
            dataRowIndex++;
        }

        if (options.AutoFitColumns || properties.Any(property => property.Width is not null))
        {
            worksheet.InsertBefore(XlsxWorksheetLayout.ConfiguredColumns(properties, columnWidths, options.AutoFitColumns), sheetData);
        }

        XlsxWorksheetLayout.AddSheet(workbookPart, worksheetPart, schema?.SheetName ?? ExcelSheetName.Default);
        ValidateLayoutConflicts(layout, startRow, startColumn, properties.Count, dataRowIndex + 1, schema?.Table is not null);
        hasFormulaCells |= XlsxWorksheetLayoutWriter.Apply(worksheetPart, worksheet, styles, layout, new HashSet<int>(properties.Select((property, index) => new { property, index }).Where(item => item.property.Width is not null).Select(item => startColumn + item.index)));
        return new WorksheetValidationContext(worksheetPart, worksheet, properties, hasFormulaCells, dataRowIndex, schema?.Table, schema?.SheetName ?? ExcelSheetName.Default, settings, startRow, startColumn, layout);
    }

    internal static WorksheetValidationContext WriteTemplate<T>(
        WorkbookPart workbookPart,
        WorkbookStyleCatalog styles,
        ExcelSchema<T> schema,
        ExcelWriteOptions options,
        ExcelSchemaOverlay<T>? overlay = null,
        CancellationToken cancellationToken = default,
        WorksheetLayoutDefinition? layout = null)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var properties = ExportProperty.For(schema, overlay);
        ValidateWorksheetShape(properties);
        var settings = schema.WorksheetSettings.Resolve(options);
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData();
        var worksheet = new Worksheet();
        XlsxWorksheetLayout.AddFrozenHeaderView(worksheet, settings);
        worksheet.AppendChild(XlsxWorksheetLayout.TemplateColumns(
            properties,
            properties.Select(property => CellValueWriter.DisplayWidth(property.Header)).ToArray(),
            options.AutoFitColumns,
            styles,
            options));
        worksheet.AppendChild(sheetData);
        var dataStart = layout?.DataStart;
        var startRow = dataStart?.FromRow ?? 1;
        var startColumn = dataStart?.FromColumn ?? 1;
        sheetData.AppendChild(HeaderRow(properties, styles.HeaderStyleIndex(options), startRow, startColumn));
        worksheetPart.Worksheet = worksheet;

        XlsxWorksheetLayout.AddSheet(workbookPart, worksheetPart, schema.SheetName);
        ValidateLayoutConflicts(layout, startRow, startColumn, properties.Count, 1, schema.Table is not null);
        var hasFormulaCells = XlsxWorksheetLayoutWriter.Apply(worksheetPart, worksheet, styles, layout, new HashSet<int>(properties.Select((property, index) => new { property, index }).Where(item => item.property.Width is not null).Select(item => startColumn + item.index)));
        return new WorksheetValidationContext(worksheetPart, worksheet, properties, hasFormulaCells, 0, schema.Table, schema.SheetName, settings, startRow, startColumn, layout);
    }

    internal static WorksheetValidationContext WriteLayout(WorkbookPart workbookPart, WorkbookStyleCatalog styles, WorksheetLayoutDefinition layout, CancellationToken cancellationToken)
    {
        cancellationToken.ThrowIfCancellationRequested();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var worksheet = new Worksheet(new SheetData());
        worksheetPart.Worksheet = worksheet;
        var hasFormulaCells = XlsxWorksheetLayoutWriter.Apply(worksheetPart, worksheet, styles, layout);
        XlsxWorksheetLayout.AddSheet(workbookPart, worksheetPart, layout.Name);
        return new WorksheetValidationContext(worksheetPart, worksheet, Array.Empty<ExportProperty>(), hasFormulaCells, 0, null, layout.Name, WorksheetSettingsDefinition.From(ExcelWriteOptions.Default), 1, 1, layout);
    }

    private static Row HeaderRow(IEnumerable<ExportProperty> properties, uint styleIndex, int rowNumber, int startColumn)
    {
        var row = XlsxWorksheetLayout.HeaderRow(properties, styleIndex);
        row.RowIndex = (uint)rowNumber;
        var cells = row.Elements<Cell>().ToArray();
        for (var index = 0; index < cells.Length; index++) cells[index].CellReference = ExcelRangeReference.ColumnName(startColumn + index) + rowNumber;
        return row;
    }

    private static void ValidateLayoutConflicts(WorksheetLayoutDefinition? layout, int startRow, int startColumn, int columns, int rows, bool hasTable)
    {
        if (layout is null || columns == 0) return;
        var dataRegion = new ExcelRangeReference(startRow, startColumn, startRow + rows - 1, startColumn + columns - 1);
        foreach (var merge in layout.Merges)
            if (merge.Overlaps(dataRegion)) throw new InvalidOperationException($"Range {merge} overlaps the schema data range {dataRegion}.");
        if (hasTable && layout.Values.Any(value => value.Range.Overlaps(dataRegion)))
            throw new InvalidOperationException($"A layout value overlaps the table range {dataRegion}.");
    }

    private static void ValidateWorksheetShape(IReadOnlyCollection<ExportProperty> properties)
    {
        if (properties.Count > MaximumColumns)
        {
            throw new InvalidOperationException($"An XLSX worksheet cannot contain more than {MaximumColumns} columns.");
        }
    }

    private static string Formula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula))
        {
            throw new InvalidOperationException("A formula must not be empty.");
        }

        var value = formula.Trim();
        if (value.StartsWith("==", StringComparison.Ordinal))
        {
            throw new InvalidOperationException("A formula can have at most one leading '='.");
        }

        return value[0] == '=' ? value.Substring(1) : value;
    }
}

internal sealed class WorksheetValidationContext
{
    internal WorksheetValidationContext(
        WorksheetPart worksheetPart,
        Worksheet worksheet,
        IReadOnlyList<ExportProperty> properties,
        bool hasFormulaCells,
        int dataRowCount,
        SchemaTableDefinition? table,
        string sheetName,
        ResolvedWorksheetSettings settings,
        int dataStartRow = 1,
        int dataStartColumn = 1,
        WorksheetLayoutDefinition? layout = null)
    {
        WorksheetPart = worksheetPart;
        Worksheet = worksheet;
        Properties = properties;
        HasFormulaCells = hasFormulaCells;
        DataRowCount = dataRowCount;
        Table = table;
        SheetName = sheetName;
        Settings = settings;
        DataStartRow = dataStartRow;
        DataStartColumn = dataStartColumn;
        Layout = layout;
    }

    internal WorksheetPart WorksheetPart { get; }

    internal Worksheet Worksheet { get; }

    internal IReadOnlyList<ExportProperty> Properties { get; }

    internal bool HasFormulaCells { get; }

    internal int DataRowCount { get; }

    internal SchemaTableDefinition? Table { get; }

    internal string SheetName { get; }

    internal ResolvedWorksheetSettings Settings { get; }
    internal int DataStartRow { get; }
    internal int DataStartColumn { get; }
    internal WorksheetLayoutDefinition? Layout { get; }
}
