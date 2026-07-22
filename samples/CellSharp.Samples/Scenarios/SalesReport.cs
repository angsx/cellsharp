using CellSharp;

namespace CellSharp.Samples.Scenarios;

internal static class SalesReport
{
    private const string Logo = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/ScL0XQAAAABJRU5ErkJggg==";

    internal static void Run(string outputDirectory)
    {
        var rows = new[]
        {
            new Sale("Acme", 12000m, "OK", new DateTime(2026, 7, 3)),
            new Sale("Contoso", 8500m, "Review", new DateTime(2026, 7, 11)),
            new Sale("Northwind", 15200m, "OK", new DateTime(2026, 7, 21)),
        };
        var schema = Excel.Schema<Sale>()
            .SheetName("Sales")
            .AsTable("SalesDetail")
            .Column(row => row.Customer)
            .Column(row => row.Revenue, column => column.Format("€ #,##0"))
            .Column(row => row.Status, column => column.AllowedValues("OK", "Review"))
            .Column(row => row.OrderDate, column => column.Format("dd MMM yyyy"))
            .Build();

        using var logo = new MemoryStream(Convert.FromBase64String(Logo));
        Excel.Workbook().AddSheet(rows, schema, sheet =>
        {
            sheet.DataStartAt("A10");
            sheet.AddImage(logo, ExcelImageFormat.Png).At("A1").Size(48, 48).Name("CellSharp logo");
            sheet.Title("Sales Performance", "B1:F2");
            sheet.Note("Report period: 1–31 July 2026. Preliminary figures; review status before distribution.", "B3:F3");

            var revenue = sheet.Kpi("Revenue", null, "A5:B7");
            revenue.Value.Formula("SUM(RevenueData)").NumberFormat("€ #,##0");
            sheet.Kpi("Orders", rows.Length, "C5:D7");
            var margin = sheet.Kpi("Margin", 0.173m, "E5:F7");
            margin.Value.NumberFormat("0.0%").ConditionalFormat().LessThan(0.15).Style(style => style.FillColor("#FFC7CE"));

            sheet.Section("Order detail", "A9:F9");
            sheet.Range("B11:B200").Name("RevenueData");
            sheet.Cell("B3").Comment("Source: ERP daily export.").Hyperlink("https://example.com/sales");
            sheet.Range("B11:B200").ConditionalFormat().GreaterThan(10000).Style(style => style.FillColor("#C6EFCE").Bold());
            sheet.Footer.Right("Page &P of &N");
            sheet.Row(9).PageBreakAfter();
            sheet.Columns("A:D").AutoFit();
        }).Write(Path.Combine(outputDirectory, "SalesReport.xlsx"));
    }

    private sealed record Sale(string Customer, decimal Revenue, string Status, DateTime OrderDate);
}
