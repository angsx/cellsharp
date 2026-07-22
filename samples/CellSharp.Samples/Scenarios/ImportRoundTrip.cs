using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class ImportRoundTrip
{
    internal static void Run(string outputDirectory)
    {
        var path = Path.Combine(outputDirectory, "product-import-round-trip.xlsx");
        var products = new[]
        {
            new ProductImportRow { Name = "USB-C Dock", Category = "Hardware", Quantity = 24, AvailableFrom = new DateTime(2025, 6, 1), Notes = "Standard configuration" },
            new ProductImportRow { Name = "Support plan", Category = "Service", Quantity = 10, AvailableFrom = new DateTime(2025, 6, 15) },
        };

        Excel.Write(path, products, DataEntryTemplate.Schema, options => options.FreezeHeaderRow());
        var result = Excel.Read(path, DataEntryTemplate.Schema, new ExcelReadOptions(ExcelInvalidRowPolicy.Include));

        Console.WriteLine($"  Imported {result.Rows.Count} product row(s); {result.Errors.Count} error(s).");
        foreach (var error in result.Errors)
        {
            Console.WriteLine($"  {error.CellReference}: {error.Message}");
        }
    }
}
