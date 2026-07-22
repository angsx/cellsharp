using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class AttributeSchema
{
    internal static void Run(string outputDirectory)
    {
        var customers = new[]
        {
            new AttributedCustomer { Id = 201, Name = "Adventure Works", CreatedAt = new DateTime(2025, 4, 3), InternalCode = "AW-01" },
            new AttributedCustomer { Id = 202, Name = "Fabrikam", CreatedAt = new DateTime(2025, 4, 17), InternalCode = "FA-02" },
        };

        var schema = Excel.SchemaFromAttributes<AttributedCustomer>(builder => builder
            .SheetName("Customers")
            .Column(x => x.Name, column => column.Header("Customer name").Width(28)));

        Excel.Write(Path.Combine(outputDirectory, "attribute-schema.xlsx"), customers, schema, options => options.FreezeHeaderRow());
    }
}
