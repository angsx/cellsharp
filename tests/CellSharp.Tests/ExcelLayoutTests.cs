using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelLayoutTests
{
    [Fact]
    public void LayoutOnlySheetWritesNativeCellsMergesStylesAndHyperlinks()
    {
        var path = Path.Combine(Path.GetTempPath(), $"cellsharp-layout-{Guid.NewGuid():N}.xlsx");
        try
        {
            Excel.Workbook().AddSheet("Report", sheet =>
            {
                sheet.Merge("A1:D1").Value("Sales report").Style(style => style.Bold().FontSize(18).AlignCenter().FillColor("#1F4E78").FontColor("#FFFFFF"));
                sheet.Range("A3:A4").Merge().Value("EU").Style(style => style.VerticalText().AlignCenter().VerticalAlignCenter().Border(border => border.Outline(ExcelBorderStyle.Thin, "#000000")));
                sheet.Cell("D5").Formula("=SUM(B5:C5)").NumberFormat("#,##0.00");
                sheet.Cell("A6").Value("OpenAI").Hyperlink("https://openai.com");
                sheet.Row(1).Height(30); sheet.Column("D").Width(20); sheet.Column("C").Hidden();
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheetPart = document.WorkbookPart!.WorksheetParts.Single();
            var worksheet = worksheetPart.Worksheet!;
            Assert.Equal(["A1:D1", "A3:A4"], worksheet.GetFirstChild<MergeCells>()!.Elements<MergeCell>().Select(value => value.Reference!.Value));
            var title = Cell(worksheet, "A1");
            Assert.Equal("Sales report", title.InlineString!.Text!.Text);
            Assert.NotNull(Cell(worksheet, "D5").CellFormula);
            Assert.Single(worksheet.GetFirstChild<Hyperlinks>()!.Elements<Hyperlink>());
            Assert.Equal(30D, worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First().Height!.Value);
            Assert.Contains(worksheet.GetFirstChild<Columns>()!.Elements<Column>(), column => column.Min!.Value == 3U && column.Hidden!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { if (File.Exists(path)) File.Delete(path); }
    }

    [Fact]
    public void LayoutRejectsInvalidRangesAndOverlappingMerges()
    {
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Report", sheet => sheet.Range("B2:A1")));
        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Workbook().AddSheet("Report", sheet => sheet.Cell("XFE1")));
        Assert.Throws<InvalidOperationException>(() => Excel.Workbook().AddSheet("Report", sheet => { sheet.Merge("A1:C3"); sheet.Merge("B2:D4"); }));
    }

    private static Cell Cell(Worksheet worksheet, string reference) => worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>().Single(cell => cell.CellReference!.Value == reference);
}
