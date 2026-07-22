using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class Report
{
    internal static void Run(string outputDirectory)
    {
        var report = new[]
        {
            new ReportRow { Customer = "Acme Ltd", Orders = 12, Revenue = 18450.75m, ReportedAt = new DateTime(2025, 6, 30) },
            new ReportRow { Customer = "Contoso", Orders = 8, Revenue = 9260.00m, ReportedAt = new DateTime(2025, 6, 30) },
        };
        var schema = Excel.Schema<ReportRow>()
            .SheetName("June report")
            .AsTable("JuneReport", "TableStyleMedium2")
            .FreezePanes(1, 0)
            .Landscape()
            .FitToPage(1, 0)
            .RepeatHeaderRowOnPrint()
            .Column(x => x.Customer, column => column.Width(28))
            .Column(x => x.Orders, column => column.Format("0").Width(10))
            .Column(x => x.Revenue, column => column.Format("#,##0.00").Width(16))
            .Column(x => x.ReportedAt, column => column.Header("As of").Format("yyyy-mm-dd").Width(14))
            .Build();

        Excel.Write(Path.Combine(outputDirectory, "sales-report.xlsx"), report, schema, options => options.AlternatingRows());
    }
}
