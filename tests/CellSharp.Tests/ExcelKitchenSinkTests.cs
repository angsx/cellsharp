using System.Buffers.Binary;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelKitchenSinkTests
{
    [Fact]
    public void CombinedMultiSheetWorkbookUsesValidWorksheetOrdering()
    {
        var path = TemporaryPath();
        var sales = Excel.Schema<Sale>()
            .SheetName("Sales Report")
            .AsTable("SalesTable")
            .Landscape()
            .FitToPage(1, 0)
            .RepeatHeaderRowOnPrint()
            .PrintGridlines()
            .Column(row => row.Customer)
            .Column(row => row.Revenue, column => column.Format("€ #,##0"))
            .Column(row => row.Status, column => column.AllowedValues("OK", "Review"))
            .Build();
        var input = Excel.Schema<Input>()
            .SheetName("John's Data")
            .AutoFilter()
            .FreezePanes(1, 1)
            .Portrait()
            .FitToPage(1, 0)
            .PrintGridlines()
            .Column(row => row.Category, column => column.AllowedValues("A", "B"))
            .Column(row => row.Amount, column => column.Formula(_ => "'Sales Report'!C13").Format("€ #,##0"))
            .Build();

        try
        {
            Excel.Workbook()
                .AddSheet([new Sale("Acme", 12000m, "OK"), new Sale("Contoso", 8500m, "Review")], sales, sheet =>
                {
                    sheet.DataStartAt("B12");
                    sheet.AddImage(new MemoryStream(Png()), ExcelImageFormat.Png).At("A1").Size(48, 48);
                    sheet.Title("Sales Performance", "B1:G2");
                    sheet.Note("Report period: July. Preliminary figures.", "B3:G3");
                    sheet.Kpi("Revenue", null, "A5:B7").Value.Formula("SUM(RevenueValues)").NumberFormat("€ #,##0");
                    sheet.Kpi("Orders", 2, "C5:D7");
                    sheet.Kpi("Margin", 0.17m, "E5:F7").Value.NumberFormat("0.0%").ConditionalFormat().LessThan(0.15).Style(style => style.FillColor("#FFC7CE"));
                    sheet.Section("Order detail", "B10:G10");
                    sheet.Range("C13:C100").Name("RevenueValues").ConditionalFormat().GreaterThan(10000).Style(style => style.FillColor("#C6EFCE"));
                    sheet.Cell("B3").Comment("Source: ERP", "Finance").Hyperlink("https://example.com/sales");
                    sheet.Columns("B:D").AutoFit();
                    sheet.Row(15).Hidden();
                    sheet.Column("G").Hidden().PageBreakAfter();
                    sheet.Footer.Right("Page &P of &N");
                    sheet.Row(20).PageBreakAfter();
                })
                .AddSheet([new Input("A", 0m)], input, sheet =>
                {
                    sheet.DataStartAt("A5");
                    sheet.AddImage(new MemoryStream(Png()), ExcelImageFormat.Png).At("E1").Width(24);
                    sheet.Title("Input review", "A1:D1");
                    sheet.Note("Values are checked before publication.", "A2:D3");
                    sheet.Cell("A2").Comment("Owned by Operations", "Operations").Hyperlink("https://example.com/input");
                    sheet.Range("B6:B100").ConditionalFormat().GreaterThan(0).Style(style => style.Bold());
                    sheet.Header.Center("Input review");
                    sheet.Footer.Left("Internal");
                    sheet.Row(8).PageBreakAfter();
                })
                .AddSheet("Cover", sheet =>
                {
                    sheet.Title("Workbook cover", "A1:D2");
                    sheet.Cell("XFD1048576").Value("Bottom");
                })
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Empty(new OpenXmlValidator().Validate(document));
            Assert.Equal(4, document.WorkbookPart!.Workbook!.Sheets!.Count()); // Includes the hidden validation lookup worksheet.
            Assert.Equal("'Sales Report'!$C$13:$C$100", document.WorkbookPart.Workbook.DefinedNames!.Elements<DefinedName>().Single(name => name.Name!.Value == "RevenueValues").Text);

            foreach (var part in document.WorkbookPart.WorksheetParts)
            {
                AssertSchemaOrder(part.Worksheet!);
            }

            var inputPart = document.WorkbookPart.WorksheetParts.Single(part => part.Worksheet!.GetFirstChild<AutoFilter>() is not null);
            Assert.NotNull(inputPart.Worksheet!.GetFirstChild<DataValidations>());
            Assert.NotNull(inputPart.Worksheet.GetFirstChild<Hyperlinks>());
            Assert.NotNull(inputPart.Worksheet.GetFirstChild<Drawing>());
            Assert.NotNull(inputPart.Worksheet.GetFirstChild<LegacyDrawing>());

            var cover = document.WorkbookPart.WorksheetParts.Single(part => part.Worksheet!.Descendants<Cell>().Any(cell => cell.CellReference!.Value == "XFD1048576"));
            Assert.Equal(3, cover.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Count());
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static void AssertSchemaOrder(Worksheet worksheet)
    {
        var order = -1;
        foreach (var element in worksheet.ChildElements)
        {
            var next = element switch
            {
                SheetProperties => 0,
                SheetViews => 1,
                Columns => 2,
                SheetData => 3,
                AutoFilter => 4,
                MergeCells => 5,
                ConditionalFormatting => 6,
                DataValidations => 7,
                Hyperlinks => 8,
                PrintOptions => 9,
                PageMargins => 10,
                PageSetup => 11,
                HeaderFooter => 12,
                RowBreaks => 13,
                ColumnBreaks => 14,
                Drawing => 15,
                LegacyDrawing => 16,
                TableParts => 17,
                _ => order,
            };
            Assert.True(next >= order, $"Worksheet element {element.GetType().Name} is out of schema order.");
            order = next;
        }
    }

    private static byte[] Png()
    {
        var bytes = new byte[24];
        new byte[] { 137, 80, 78, 71, 13, 10, 26, 10 }.CopyTo(bytes, 0);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(16, 4), 1);
        BinaryPrimitives.WriteInt32BigEndian(bytes.AsSpan(20, 4), 1);
        return bytes;
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-kitchen-sink-{Guid.NewGuid():N}.xlsx");
    private sealed record Sale(string Customer, decimal Revenue, string Status);
    private sealed record Input(string Category, decimal Amount);
}
