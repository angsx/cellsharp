using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class FluentSchema
{
    internal static void Run(string outputDirectory)
    {
        var customers = new[]
        {
            new Customer { Id = 101, Name = "Northwind Traders", CreatedAt = new DateTime(2025, 3, 1) },
            new Customer { Id = 102, Name = "Tailspin Toys", CreatedAt = new DateTime(2025, 3, 12) },
        };

        var schema = Excel.Schema<Customer>()
            .SheetName("Customers")
            .Column(x => x.Id, column => column.Header("Customer ID").Width(14))
            .Column(x => x.Name, column => column.Header("Company").Width(30))
            .Column(x => x.CreatedAt, column => column.Header("Created").Format("yyyy-mm-dd").Width(14))
            .Build();

        Excel.Write(Path.Combine(outputDirectory, "fluent-schema.xlsx"), customers, schema, options => options.FreezeHeaderRow());
    }
}
