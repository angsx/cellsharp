using CellSharp;

namespace CellSharp.Samples.Scenarios;

internal static class ConditionalFormatting
{
    internal static void Run(string outputDirectory)
    {
        var rows = new[]
        {
            new Order(1001, "Acme Ltd", 18500m, .31m, DateTime.Today.AddDays(10), "OK"),
            new Order(1002, "Contoso", 750m, -.04m, DateTime.Today.AddDays(-2), "ERROR"),
            new Order(1002, "Northwind", 4200m, .18m, DateTime.Today.AddDays(4), "WARN: review"),
        };
        var schema = Excel.Schema<Order>()
            .SheetName("Orders")
            .AsTable("OrdersTable")
            .Column(order => order.OrderId)
            .Column(order => order.Customer)
            .Column(order => order.Revenue, column => column.Format("€ #,##0.00"))
            .Column(order => order.Margin, column => column.Format("0.0%"))
            .Column(order => order.DueDate, column => column.Format("yyyy-mm-dd"))
            .Column(order => order.Status, column => column.AllowedValues("OK", "ERROR", "WARN: review"))
            .Build();

        Excel.Workbook().AddSheet(rows, schema, sheet =>
        {
            sheet.DataStartAt("A4");
            sheet.Merge("A1:F1").Value("Order health report").Style(s => s.Bold().FontSize(18).FillColor("#1F4E78").FontColor("#FFFFFF").AlignCenter());
            sheet.Range("C5:C200").ConditionalFormat().GreaterThan(1000m).Style(s => s.FillColor("#C6EFCE").FontColor("#006100"));
            sheet.Range("D5:D200").ConditionalFormat().LessThan(0).Style(s => s.FillColor("#FFC7CE").FontColor("#9C0006")).StopIfTrue();
            sheet.Range("E5:E200").ConditionalFormat().Formula("E5<TODAY()").Style(s => s.FontColor("#9C0006").Bold());
            sheet.Range("A5:A200").ConditionalFormat().DuplicateValues().Style(s => s.FillColor("#FFEB9C"));
            sheet.Range("F5:F200").ConditionalFormat().ContainsText("ERROR").Style(s => s.FillColor("#FFC7CE").Bold());
            sheet.Columns("A:F").AutoFit();
        }).Write(Path.Combine(outputDirectory, "ConditionalFormatting.xlsx"));
    }

    private sealed record Order(int OrderId, string Customer, decimal Revenue, decimal Margin, DateTime DueDate, string Status);
}
