using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxWorksheetLayoutWriter
{
    internal static bool Apply(WorksheetPart part, Worksheet worksheet, WorkbookStyleCatalog styles, WorksheetLayoutDefinition? layout, ISet<int>? schemaWidthColumns = null)
    {
        if (layout is null) return false;
        var data = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
        var cells = Index(data);
        foreach (var operation in layout.Values)
        {
            foreach (var coordinate in Coordinates(operation.Range))
            {
                var cell = GetCell(data, cells, coordinate.row, coordinate.column);
                var styleIndex = cell.StyleIndex?.Value;
                var replacement = operation.Formula is null ? CellValueWriter.Create(operation.Value, styleIndex) : CellValueWriter.CreateFormula(operation.Formula, null, styleIndex);
                replacement.CellReference = cell.CellReference;
                cell.InsertBeforeSelf(replacement); cell.Remove(); cells[(coordinate.row, coordinate.column)] = replacement;
            }
        }
        var effectiveStyles = new Dictionary<(int row, int column), LayoutStyleDefinition>();
        var styleTargets = new HashSet<(int row, int column)>(cells.Keys);
        foreach (var operation in layout.Styles)
            foreach (var coordinate in Coordinates(operation.Range)) styleTargets.Add(coordinate);
        // Column and row styles affect materialized cells plus cells requested by a range style.
        // This keeps a column-only operation sparse while retaining column -> row -> range -> cell precedence.
        foreach (var operation in layout.ColumnStyles)
            foreach (var coordinate in styleTargets.Where(cell => cell.column >= operation.FromColumn && cell.column <= operation.ToColumn).ToArray())
                Overlay(effectiveStyles, coordinate, operation.Style);
        foreach (var operation in layout.RowStyles)
            foreach (var coordinate in styleTargets.Where(cell => cell.row == operation.Row).ToArray())
                Overlay(effectiveStyles, coordinate, operation.Style);
        foreach (var operation in layout.Styles.Where(operation => !operation.Range.IsCell))
            foreach (var coordinate in Coordinates(operation.Range))
                Overlay(effectiveStyles, coordinate, MaterializeInsideBorders(operation.Style, operation.Range, coordinate.row, coordinate.column));
        foreach (var operation in layout.Styles.Where(operation => operation.Range.IsCell))
            Overlay(effectiveStyles, (operation.Range.FromRow, operation.Range.FromColumn), operation.Style);
        foreach (var item in effectiveStyles)
        {
            var cell = GetCell(data, cells, item.Key.row, item.Key.column);
            cell.StyleIndex = styles.ComposeStyleIndex(cell.StyleIndex?.Value, item.Value);
        }
        foreach (var merge in layout.Merges)
        {
            var merges = worksheet.GetFirstChild<MergeCells>();
            if (merges is null) { merges = new MergeCells(); worksheet.AppendChild(merges); }
            merges.AppendChild(new MergeCell { Reference = merge.ToString() }); merges.Count = (uint)merges.Elements<MergeCell>().Count();
        }
        XlsxConditionalFormattingWriter.Apply(worksheet, styles, layout.ConditionalFormats);
        ApplyRows(data, layout.Rows);
        ApplyColumns(worksheet, layout.Columns);
        ApplyAutoFitColumns(worksheet, data, layout.Columns, schemaWidthColumns);
        ApplyHyperlinks(part, worksheet, layout.Hyperlinks);
        return layout.Values.Any(value => value.Formula is not null);
    }

    private static void Overlay(Dictionary<(int row, int column), LayoutStyleDefinition> styles, (int row, int column) coordinate, LayoutStyleDefinition value) =>
        styles[coordinate] = styles.TryGetValue(coordinate, out var current) ? current.Overlay(value) : value;

    private static Dictionary<(int row, int column), Cell> Index(SheetData data)
    {
        var result = new Dictionary<(int row, int column), Cell>();
        foreach (var row in data.Elements<Row>())
        {
            var rowIndex = (int)(row.RowIndex?.Value ?? 0U); var column = 0;
            foreach (var cell in row.Elements<Cell>()) { column++; if (cell.CellReference?.Value is { } reference) { var parsed = ExcelRangeReference.ParseCell(reference); result[(parsed.FromRow, parsed.FromColumn)] = cell; } else if (rowIndex > 0) { cell.CellReference = ExcelRangeReference.ColumnName(column) + rowIndex; result[(rowIndex, column)] = cell; } }
        }
        return result;
    }
    private static Cell GetCell(SheetData data, Dictionary<(int row, int column), Cell> cells, int rowNumber, int columnNumber)
    {
        if (cells.TryGetValue((rowNumber, columnNumber), out var found)) return found;
        var row = data.Elements<Row>().FirstOrDefault(value => value.RowIndex?.Value == (uint)rowNumber);
        if (row is null) { row = new Row { RowIndex = (uint)rowNumber }; var next = data.Elements<Row>().FirstOrDefault(value => (value.RowIndex?.Value ?? 0U) > (uint)rowNumber); if (next is null) data.AppendChild(row); else data.InsertBefore(row, next); }
        var cell = new Cell { CellReference = ExcelRangeReference.ColumnName(columnNumber) + rowNumber };
        var nextCell = row.Elements<Cell>().FirstOrDefault(value => ExcelRangeReference.ParseCell(value.CellReference!.Value!).FromColumn > columnNumber);
        if (nextCell is null) row.AppendChild(cell); else row.InsertBefore(cell, nextCell);
        cells[(rowNumber, columnNumber)] = cell; return cell;
    }
    private static IEnumerable<(int row, int column)> Coordinates(ExcelRangeReference range)
    { for (var row = range.FromRow; row <= range.ToRow; row++) for (var column = range.FromColumn; column <= range.ToColumn; column++) yield return (row, column); }
    private static LayoutStyleDefinition MaterializeInsideBorders(LayoutStyleDefinition style, ExcelRangeReference range, int row, int column)
    {
        if (style.Border is null) return style;
        var border = style.Border;
        if (column < range.ToColumn && border.InsideVertical is not null) border = border with { Right = border.Right ?? border.InsideVertical };
        if (row < range.ToRow && border.InsideHorizontal is not null) border = border with { Bottom = border.Bottom ?? border.InsideHorizontal };
        return style with { Border = border };
    }
    private static void ApplyRows(SheetData data, IReadOnlyDictionary<int, RowLayoutDefinition> rows)
    { foreach (var item in rows) { var row = data.Elements<Row>().FirstOrDefault(value => value.RowIndex?.Value == (uint)item.Key); if (row is null) { row = new Row { RowIndex = (uint)item.Key }; data.AppendChild(row); } if (item.Value.Height is not null) { row.Height = item.Value.Height; row.CustomHeight = true; } if (item.Value.Hidden is not null) row.Hidden = item.Value.Hidden; } }
    private static void ApplyColumns(Worksheet worksheet, IReadOnlyDictionary<int, ColumnLayoutDefinition> columns)
    { if (columns.Count == 0) return; var collection = Columns(worksheet); foreach (var item in columns) { var column = Column(collection, item.Key); if (item.Value.Width is not null) { column.Width = item.Value.Width; column.CustomWidth = true; } if (item.Value.Hidden is not null) column.Hidden = item.Value.Hidden; } }
    private static void ApplyAutoFitColumns(Worksheet worksheet, SheetData data, IReadOnlyDictionary<int, ColumnLayoutDefinition> columns, ISet<int>? schemaWidthColumns)
    {
        var requested = columns.Where(item => item.Value.AutoFit == true && item.Value.Width is null && (schemaWidthColumns is null || !schemaWidthColumns.Contains(item.Key))).Select(item => item.Key).ToArray();
        if (requested.Length == 0) return;
        var widths = requested.ToDictionary(column => column, _ => 0);
        foreach (var cell in data.Descendants<Cell>())
        {
            if (cell.CellReference?.Value is not { } reference) continue;
            var column = ExcelRangeReference.ParseCell(reference).FromColumn;
            if (!widths.ContainsKey(column)) continue;
            var text = cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;
            widths[column] = Math.Max(widths[column], CellValueWriter.DisplayWidth(text));
        }
        var collection = Columns(worksheet);
        foreach (var item in widths) { var column = Column(collection, item.Key); column.Width = Math.Max(10D, Math.Min(60D, item.Value + 2D)); column.CustomWidth = true; }
    }
    private static Columns Columns(Worksheet worksheet)
    { var collection = worksheet.GetFirstChild<Columns>(); if (collection is not null) return collection; collection = new Columns(); var data = worksheet.GetFirstChild<SheetData>(); if (data is null) worksheet.AppendChild(collection); else worksheet.InsertBefore(collection, data); return collection; }
    private static Column Column(Columns columns, int index) => columns.Elements<Column>().FirstOrDefault(value => value.Min?.Value == (uint)index && value.Max?.Value == (uint)index) ?? columns.AppendChild(new Column { Min = (uint)index, Max = (uint)index });
    private static void ApplyHyperlinks(WorksheetPart part, Worksheet worksheet, IReadOnlyList<HyperlinkOperation> links)
    { if (links.Count == 0) return; var hyperlinks = worksheet.GetFirstChild<Hyperlinks>() ?? worksheet.AppendChild(new Hyperlinks()); foreach (var link in links) foreach (var coordinate in Coordinates(link.Range)) { var reference = ExcelRangeReference.ColumnName(coordinate.column) + coordinate.row; var hyperlink = new Hyperlink { Reference = reference }; if (link.Target.StartsWith("#", StringComparison.Ordinal)) hyperlink.Location = link.Target.Substring(1); else hyperlink.Id = part.AddHyperlinkRelationship(new Uri(link.Target, UriKind.Absolute), true).Id; hyperlinks.AppendChild(hyperlink); } }
}
