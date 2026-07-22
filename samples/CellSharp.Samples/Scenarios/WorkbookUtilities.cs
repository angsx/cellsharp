using CellSharp;

namespace CellSharp.Samples.Scenarios;

internal static class WorkbookUtilities
{
    internal static void Run(string outputDirectory)
    {
        var sales = new[]
        {
            new Sale("North", 12000m, "OK"),
            new Sale("South", 8500m, "Review"),
            new Sale("West", 15200m, "OK"),
        };
        var costs = new[]
        {
            new Cost("Hosting", 3200m),
            new Cost("Travel", 900m),
        };
        var salesSchema = Excel.Schema<Sale>()
            .SheetName("Sales")
            .AsTable("SalesTable")
            .Landscape()
            .FitToPage(1, 0)
            .Column(row => row.Region)
            .Column(row => row.Revenue, column => column.Format("€ #,##0"))
            .Column(row => row.Status, column => column.AllowedValues("OK", "Review"))
            .Build();
        var costsSchema = Excel.Schema<Cost>()
            .SheetName("Costs")
            .AsTable("CostsTable")
            .Column(row => row.Category)
            .Column(row => row.Amount, column => column.Format("€ #,##0"))
            .Build();

        Excel.Workbook()
            .AddSheet(sales, salesSchema, sheet =>
            {
                sheet.DataStartAt("A4");
                sheet.Merge("A1:C1").Value("Sales report").Style(s => s.Bold().FontSize(18).FillColor("#1F4E78").FontColor("#FFFFFF").AlignCenter());
                sheet.Range("B5:B100").Name("Revenue");
                sheet.Cell("B8").Formula("SUM(Revenue)").Comment("Total uses the workbook-scoped Revenue name.");
                sheet.Range("B5:B100").ConditionalFormat().GreaterThan(10000).Style(s => s.FillColor("#C6EFCE").Bold());
                sheet.Header.Left("CellSharp");
                sheet.Header.Center("Sales report");
                sheet.Footer.Right("Page &P of &N");
                sheet.Row(8).PageBreakAfter();
                sheet.Columns("A:C").AutoFit();
            })
            .AddSheet(costs, costsSchema, sheet =>
            {
                sheet.Range("B2:B100").Name("Costs");
                sheet.Cell("A2").Comment("Costs imported from Finance.", "Finance");
                sheet.Header.Center("Costs report");
                sheet.Footer.Left("Confidential");
                sheet.Column("B").PageBreakAfter();
            })
            .Write(Path.Combine(outputDirectory, "WorkbookUtilities.xlsx"));
    }

    private sealed record Sale(string Region, decimal Revenue, string Status);
    private sealed record Cost(string Category, decimal Amount);
}
