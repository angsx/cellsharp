using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelHeaderFooterAndPageBreakTests
{
    [Fact]
    public void HeadersFootersAndPageBreaksAreNativeAndCoexistWithPrintSettings()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Row>()
            .SheetName("Sales")
            .Landscape()
            .FitToPage(1, 0)
            .RepeatHeaderRowOnPrint()
            .PrintGridlines()
            .Column(row => row.Value)
            .Build();
        try
        {
            Excel.Workbook().AddSheet([new Row(1)], schema, sheet =>
            {
                sheet.Header.Left("CellSharp & Co");
                sheet.Header.Center("Sales &A");
                sheet.Header.Right("&D");
                sheet.Footer.Left("Confidential");
                sheet.Footer.Right("Page &P of &N");
                sheet.Row(50).PageBreakAfter().PageBreakAfter();
                sheet.Row(100).PageBreakAfter();
                sheet.Column("F").PageBreakAfter().PageBreakAfter();
                sheet.Column("J").PageBreakAfter();
                sheet.Row(100).Hidden();
                sheet.Column("J").Hidden();
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Equal("&LCellSharp && Co&CSales &A&R&D", worksheet.GetFirstChild<HeaderFooter>()!.OddHeader!.Text);
            Assert.Equal("&LConfidential&RPage &P of &N", worksheet.GetFirstChild<HeaderFooter>()!.OddFooter!.Text);
            Assert.Equal(new uint[] { 50, 100 }, worksheet.GetFirstChild<RowBreaks>()!.Elements<Break>().Select(@break => @break.Id!.Value));
            Assert.Equal(new uint[] { 6, 10 }, worksheet.GetFirstChild<ColumnBreaks>()!.Elements<Break>().Select(@break => @break.Id!.Value));
            Assert.Equal(OrientationValues.Landscape, worksheet.GetFirstChild<PageSetup>()!.Orientation!.Value);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void PageBreaksAcceptWorksheetLimitCoordinates()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Limits", sheet =>
            {
                sheet.Row(1048576).PageBreakAfter();
                sheet.Column("XFD").PageBreakAfter();
            }).Write(path);
            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Workbook().AddSheet("Limits", sheet => sheet.Row(1048577)));
        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Workbook().AddSheet("Limits", sheet => sheet.Column("XFE")));
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-print-{Guid.NewGuid():N}.xlsx");
    private static void Delete(string path) { if (File.Exists(path)) File.Delete(path); }
    private sealed record Row(int Value);
}
