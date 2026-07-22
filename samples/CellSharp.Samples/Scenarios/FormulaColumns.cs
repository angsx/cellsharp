using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class FormulaColumns
{
    internal static void Run(string outputDirectory)
    {
        var lines = new[]
        {
            new OrderLine { Product = "USB-C Dock", Quantity = 2, UnitPrice = 129.00m, Total = 258.00m },
            new OrderLine { Product = "Laptop stand", Quantity = 3, UnitPrice = 49.50m, Total = 148.50m },
        };
        var schema = Excel.Schema<OrderLine>()
            .SheetName("Order report")
            .Column(x => x.Product, column => column.Width(24))
            .Column(x => x.Quantity, column => column.Format("0").Width(10))
            .Column(x => x.UnitPrice, column => column.Header("Unit price").Format("#,##0.00").Width(14))
            .Column(x => x.Total, column => column.Header("Total").Formula(context => $"=B{context.Row}*C{context.Row}").Format("#,##0.00").Width(14))
            .Build();

        Excel.Write(Path.Combine(outputDirectory, "formula-columns.xlsx"), lines, schema, options => options.FreezeHeaderRow());
    }
}
