using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelStyleCompositionTests
{
    [Fact]
    public void SchemaRangeRowColumnAndCellStylesComposeByProperty()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Row>()
            .Column(row => row.Name, column => column.Align(ExcelHorizontalAlignment.Left))
            .Column(row => row.Price, column => column.Format("€ #,##0.00").Align(ExcelHorizontalAlignment.Right))
            .Build();
        try
        {
            Excel.Workbook().AddSheet([new Row("Ada", 12.5M)], schema, sheet =>
            {
                sheet.Column("B").Style(style => style.FontSize(11));
                sheet.Row(2).Style(style => style.VerticalAlignCenter());
                sheet.Range("A2:B2").Style(style => style.Bold());
                sheet.Cell("B2").Style(style => style.Bold(false).FontColor("#FF0000").Border(border => border.Bottom(ExcelBorderStyle.Thin, "#0000FF")));
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var price = Cell(worksheet, "B2");
            var name = Cell(worksheet, "A2");
            var priceFormat = Format(document, price);
            var nameFormat = Format(document, name);
            var formats = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.NumberingFormats!.Elements<NumberingFormat>()
                .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode!.Value);

            Assert.Equal("€ #,##0.00", formats[priceFormat.NumberFormatId!.Value]);
            Assert.Equal(HorizontalAlignmentValues.Right, priceFormat.Alignment!.Horizontal!.Value);
            Assert.Equal(VerticalAlignmentValues.Center, priceFormat.Alignment!.Vertical!.Value);
            Assert.Null(Font(document, price).Bold);
            Assert.Equal("FFFF0000", Font(document, price).Color!.Rgb!.Value);
            Assert.NotNull(Border(document, price).BottomBorder!.Color);
            Assert.NotNull(Border(document, price).TopBorder!.Color);
            Assert.NotNull(Font(document, name).Bold);
            Assert.Equal(HorizontalAlignmentValues.Left, nameFormat.Alignment!.Horizontal!.Value);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void EquivalentComposedStylesReuseOneCellFormat()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Report", sheet =>
            {
                sheet.Range("A1:A1000").Value("value").Style(style => style.Bold().FillColor("#C6EFCE"));
                sheet.Range("B1:B1000").Style(style => style.Bold());
                sheet.Range("B1:B1000").Style(style => style.FillColor("#C6EFCE"));
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!.Descendants<Cell>().ToArray();
            var styles = cells.Select(cell => cell.StyleIndex!.Value).Distinct().ToArray();
            Assert.Single(styles);
            Assert.Equal(2, document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().Count()); // Normal + resolved format
        }
        finally { Delete(path); }
    }

    [Fact]
    public void SelectiveAutofitUsesMaterializedCellsAndExplicitWidthWins()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Report", sheet =>
            {
                sheet.Cell("B4").Value("A deliberately long title");
                sheet.Cell("C4").Value("short");
                sheet.Columns("B:C").AutoFit();
                sheet.Column("C").Width(25);
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var columns = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
            var b = columns.Single(column => column.Min!.Value == 2U);
            var c = columns.Single(column => column.Min!.Value == 3U);
            Assert.True(b.Width!.Value > 10D);
            Assert.Equal(25D, c.Width!.Value);
        }
        finally { Delete(path); }
    }

    private static Cell Cell(Worksheet worksheet, string reference) => worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>().Single(cell => cell.CellReference!.Value == reference);
    private static CellFormat Format(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
    private static Font Font(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!.Elements<Font>().ElementAt((int)Format(document, cell).FontId!.Value);
    private static Border Border(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Borders!.Elements<Border>().ElementAt((int)Format(document, cell).BorderId!.Value);
    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-composition-{Guid.NewGuid():N}.xlsx");
    private static void Delete(string path) { if (File.Exists(path)) File.Delete(path); }
    private sealed record Row(string Name, decimal Price);
}
