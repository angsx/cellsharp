using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class CustomConverter
{
    internal static void Run(string outputDirectory)
    {
        var path = Path.Combine(outputDirectory, "custom-converter.xlsx");
        var schema = Excel.Schema<CustomerWithStatus>()
            .SheetName("Customers")
            .Column(x => x.Id, column => column.Header("Customer ID"))
            .Column(x => x.Name, column => column.Header("Company").Width(28))
            .Column(x => x.Status, column => column.ConvertWith(CustomerStatusConverter.Instance))
            .Build();

        Excel.Write(path, new[]
        {
            new CustomerWithStatus { Id = 401, Name = "Fourth Coffee", Status = CustomerStatus.Active },
            new CustomerWithStatus { Id = 402, Name = "Graphic Design Institute", Status = CustomerStatus.Pending },
        }, schema);

        var result = Excel.Read<CustomerWithStatus>(path, schema);
        Console.WriteLine($"  Converted {result.Rows.Count} customer status value(s); {result.Errors.Count} error(s).");
    }
}
