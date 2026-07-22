using System.Buffers.Binary;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelReportComponentTests
{
    private static readonly string[] TextComponentMerges = ["A1:D2", "A4:D4", "F4:G4", "A6:D7"];
    private static readonly string[] KpiMerges = ["A4:B4", "A5:B6", "C4:D4", "C5:D6"];

    [Fact]
    public void TitleSectionAndNoteExpandToMergedStyledLayoutPrimitives()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Report", sheet =>
            {
                sheet.Title("Sales Report", "A1:D2").Style(style => style.FontSize(24).FontColor("#1F4E78"));
                sheet.Title("Single-cell heading", "F1");
                sheet.Section("Europe", "A4:D4").Style(style => style.FillColor("#D9EAF7"));
                sheet.Section("Asia", "F4:G4");
                sheet.Note("Generated automatically from ERP data.", "A6:D7").Style(style => style.FontColor("#666666"));
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal(TextComponentMerges, worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Select(merge => merge.Reference!.Value));
            Assert.Equal("Sales Report", Cell(worksheet, "A1").InlineString!.Text!.Text);
            Assert.Equal("Single-cell heading", Cell(worksheet, "F1").InlineString!.Text!.Text);
            Assert.Equal("Europe", Cell(worksheet, "A4").InlineString!.Text!.Text);
            Assert.Equal("Asia", Cell(worksheet, "F4").InlineString!.Text!.Text);
            Assert.Equal("Generated automatically from ERP data.", Cell(worksheet, "A6").InlineString!.Text!.Text);
            var title = Format(document, Cell(worksheet, "A1"));
            Assert.Equal(24D, Font(document, title).FontSize!.Val!.Value);
            Assert.NotNull(Font(document, title).Bold);
            Assert.Equal(HorizontalAlignmentValues.Center, title.Alignment!.Horizontal!.Value);
            Assert.True(title.Alignment.WrapText!.Value);
            var section = Format(document, Cell(worksheet, "A4"));
            Assert.Equal("D9EAF7", Background(document, section));
            Assert.Equal(VerticalAlignmentValues.Center, section.Alignment!.Vertical!.Value);
            Assert.Equal(BorderStyleValues.Thin, Border(document, section).BottomBorder!.Style!.Value);
            Assert.Equal("F2F2F2", Background(document, Format(document, Cell(worksheet, "F4"))));
            var note = Format(document, Cell(worksheet, "A6"));
            Assert.True(note.Alignment!.WrapText!.Value);
            Assert.Equal(VerticalAlignmentValues.Top, note.Alignment.Vertical!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void KpiUsesTopRowForLabelAndRemainingRowsForComposableValue()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Report", sheet =>
            {
                var revenue = sheet.Kpi("Revenue", 125400m, "A4:B6");
                revenue.Value.NumberFormat("€ #,##0").Name("TotalRevenue");
                revenue.Value.Style(style => style.FontColor("#1F4E78"));
                revenue.Value.ConditionalFormat().GreaterThan(100000).Style(style => style.Bold());

                var profit = sheet.Kpi("Profit", null, "C4:D6");
                profit.Value.Formula("SUM(TotalRevenue)").NumberFormat("€ #,##0");
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal(KpiMerges, worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Select(merge => merge.Reference!.Value));
            Assert.Equal("Revenue", Cell(worksheet, "A4").InlineString!.Text!.Text);
            Assert.Equal("125400", Cell(worksheet, "A5").CellValue!.Text);
            Assert.NotNull(Font(document, Format(document, Cell(worksheet, "A4"))).Bold);
            Assert.Equal(18D, Font(document, Format(document, Cell(worksheet, "A5"))).FontSize!.Val!.Value);
            Assert.NotNull(Cell(worksheet, "C5").CellFormula);
            Assert.Equal("'Report'!$A$5:$B$6", document.WorkbookPart.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text);
            Assert.Single(worksheet.Elements<ConditionalFormatting>());
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void ComponentsValidateTextSizeAndMergeConflicts()
    {
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Report", sheet => sheet.Title(" ", "A1")));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Report", sheet => sheet.Note(" ", "A1")));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Report", sheet => sheet.Kpi("Revenue", 1, "A1:B1")));
        Assert.Throws<InvalidOperationException>(() => Excel.Workbook().AddSheet("Report", sheet =>
        {
            sheet.Merge("A1:B2");
            sheet.Title("Conflict", "A1:C2");
        }));
    }

    [Fact]
    public void FullReportComposesWithImagesTypedDataTablesAndWorksheetFeatures()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<SalesRow>().SheetName("Sales").AsTable()
            .Column(row => row.Customer)
            .Column(row => row.Amount, column => column.Format("€ #,##0"))
            .Build();
        try
        {
            Excel.Workbook().AddSheet([new SalesRow("Acme", 1200m)], schema, sheet =>
            {
                sheet.DataStartAt("A10");
                sheet.AddImage(new MemoryStream(Png()), ExcelImageFormat.Png).At("A1").Size(48, 48);
                sheet.Title("Sales Performance", "B1:F2");
                sheet.Note("Preliminary figures.", "B3:F3");
                sheet.Cell("B3").Comment("Source: ERP").Hyperlink("https://openai.com");
                var revenue = sheet.Kpi("Revenue", 125400m, "A5:B7");
                revenue.Value.Name("Revenue").NumberFormat("€ #,##0");
                sheet.Kpi("Orders", 842, "C5:D7");
                var margin = sheet.Kpi("Margin", 0.173m, "E5:F7");
                margin.Value.NumberFormat("0.0%").ConditionalFormat().LessThan(0).Style(style => style.FillColor("#FFC7CE"));
                sheet.Section("Order detail", "A9:F9");
                sheet.Footer.Right("Page &P of &N");
                sheet.Row(9).PageBreakAfter();
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var part = document.WorkbookPart!.WorksheetParts.First();
            Assert.NotNull(part.DrawingsPart);
            Assert.Single(part.TableDefinitionParts);
            Assert.NotNull(part.WorksheetCommentsPart);
            Assert.NotNull(part.Worksheet!.GetFirstChild<HeaderFooter>());
            Assert.NotNull(part.Worksheet.GetFirstChild<RowBreaks>());
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
    }

    private static Cell Cell(Worksheet worksheet, string reference) => worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>().Single(cell => cell.CellReference!.Value == reference);
    private static CellFormat Format(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ElementAt((int)cell.StyleIndex!.Value);
    private static Font Font(SpreadsheetDocument document, CellFormat format) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!.Elements<Font>().ElementAt((int)format.FontId!.Value);
    private static string? Background(SpreadsheetDocument document, CellFormat format) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fills!.Elements<Fill>().ElementAt((int)format.FillId!.Value).PatternFill?.ForegroundColor?.Rgb?.Value?[2..];
    private static Border Border(SpreadsheetDocument document, CellFormat format) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Borders!.Elements<Border>().ElementAt((int)format.BorderId!.Value);
    private static byte[] Png()
    {
        var bytes = new byte[24]; new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }.CopyTo(bytes, 0);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16, 4), 1); BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(20, 4), 1); return bytes;
    }
    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-report-{Guid.NewGuid():N}.xlsx");
    private static void Delete(string path) { if (File.Exists(path)) File.Delete(path); }
    private sealed record SalesRow(string Customer, decimal Amount);
}
