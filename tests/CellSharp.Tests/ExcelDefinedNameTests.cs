using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelDefinedNameTests
{
    [Fact]
    public void WorkbookScopedNamesUseAbsoluteEscapedReferencesAndWorkInFormulas()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook()
                .AddSheet("Sales Report's", sheet =>
                {
                    sheet.Range("B2:B100").Name("Revenue");
                    sheet.Range("C:C").Name("AllCosts");
                    sheet.Cell("D101").Formula("SUM(Revenue)");
                    sheet.Range("B2:B100").ConditionalFormat().GreaterThan(1000);
                })
                .AddSheet("Costs", sheet => sheet.Cell("A1").Name("CostCell"))
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var names = document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().ToDictionary(name => name.Name!.Value!, name => name.Text);
            Assert.Equal("'Sales Report''s'!$B$2:$B$100", names["Revenue"]);
            Assert.Equal("'Sales Report''s'!$C:$C", names["AllCosts"]);
            Assert.Equal("'Costs'!$A$1", names["CostCell"]);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void NamesValidateImmediatelyAndDuplicatesFailBeforeSaving()
    {
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Sales", sheet => sheet.Cell("A1").Name("A1")));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Sales", sheet => sheet.Cell("A1").Name("Bad Name")));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Sales", sheet => sheet.Cell("A1").Name("_xlnm.Print_Area")));

        var path = TemporaryPath();
        try
        {
            Assert.Throws<InvalidOperationException>(() => Excel.Workbook()
                .AddSheet("Sales", sheet => sheet.Cell("A1").Name("Revenue"))
                .AddSheet("Costs", sheet => sheet.Cell("A1").Name("revenue"))
                .Write(path));
        }
        finally { Delete(path); }
    }

    [Fact]
    public void NamedRangesUseFinalCoordinatesAlongsideDataStartAndTables()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Sale>().SheetName("Sales").AsTable().Column(sale => sale.Revenue).Build();
        try
        {
            Excel.Workbook().AddSheet([new Sale(10m)], schema, sheet =>
            {
                sheet.DataStartAt("B4");
                sheet.Range("B5:B100").Name("Revenue");
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Equal("'Sales'!$B$5:$B$100", document.WorkbookPart!.Workbook!.DefinedNames!.Elements<DefinedName>().Single().Text);
            Assert.Single(document.WorkbookPart.WorksheetParts.Single().TableDefinitionParts);
            Assert.Empty(new OpenXmlValidator().Validate(document));
        }
        finally { Delete(path); }
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-name-{Guid.NewGuid():N}.xlsx");
    private static void Delete(string path) { if (File.Exists(path)) File.Delete(path); }
    private sealed record Sale(decimal Revenue);
}
