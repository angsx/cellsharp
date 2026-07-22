using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelLargeDatasetSmokeTests
{
    [Fact]
    public void WritesOneHundredThousandRowsWithoutMaterializingConditionalFormattingRange()
    {
        const int rowCount = 100000;
        var schema = Excel.Schema<LargeRow>()
            .SheetName("Large data")
            .AsTable("LargeData")
            .Column(row => row.Id)
            .Column(row => row.Customer)
            .Column(row => row.Amount, column => column.Format("#,##0.00"))
            .Build();
        var rows = Enumerable.Range(1, rowCount).Select(index => new LargeRow(index, $"Customer {index}", index / 10m));
        using var stream = new MemoryStream();

        Excel.Workbook().AddSheet(rows, schema, sheet =>
        {
            sheet.Range("C2:C100001").ConditionalFormat().GreaterThan(5000).Style(style => style.Bold());
        }).Write(stream);

        Assert.True(stream.CanWrite);
        stream.Position = 0;
        using var document = SpreadsheetDocument.Open(stream, false);
        var part = document.WorkbookPart!.WorksheetParts.Single();
        Assert.Equal(rowCount + 1, part.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Count());
        Assert.Equal("A1:C100001", part.TableDefinitionParts.Single().Table!.Reference!.Value);
        Assert.Equal("C2:C100001", part.Worksheet.Elements<ConditionalFormatting>().Single().SequenceOfReferences!.InnerText);
    }

    private sealed record LargeRow(int Id, string Customer, decimal Amount);
}
