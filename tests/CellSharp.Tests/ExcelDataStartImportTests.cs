using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelDataStartImportTests
{
    [Fact]
    public void ExplicitSchemaRoundTripsDataPlacedBelowReportContent()
    {
        var path = Path.Combine(Path.GetTempPath(), $"cellsharp-data-start-import-{Guid.NewGuid():N}.xlsx");
        var schema = Excel.Schema<Row>()
            .SheetName("Sales")
            .Column(row => row.Customer)
            .Column(row => row.Revenue)
            .Build();
        try
        {
            Excel.Workbook().AddSheet([new Row { Customer = "Acme", Revenue = 12m }, new Row { Customer = "Contoso", Revenue = 24m }], schema, sheet =>
            {
                sheet.Title("Sales performance", "B1:E2");
                sheet.Note("Figures are preliminary.", "B3:E3");
                sheet.DataStartAt("B6");
            }).Write(path);

            var result = Excel.Read(path, schema);

            Assert.True(result.IsValid);
            Assert.Equal(["Acme", "Contoso"], result.Rows.Select(row => row.Customer));
            Assert.Equal([12m, 24m], result.Rows.Select(row => row.Revenue));
        }
        finally
        {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private sealed class Row
    {
        public string Customer { get; set; } = "";

        public decimal Revenue { get; set; }
    }
}
