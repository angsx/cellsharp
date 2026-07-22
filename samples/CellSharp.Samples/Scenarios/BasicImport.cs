using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class BasicImport
{
    internal static void Run(string outputDirectory)
    {
        var path = Path.Combine(outputDirectory, "basic-import.xlsx");
        Excel.Write(path, new[]
        {
            new Customer { Id = 11, Name = "Wide World Importers", CreatedAt = new DateTime(2025, 5, 4) },
            new Customer { Id = 12, Name = "Woodgrove Bank", CreatedAt = new DateTime(2025, 5, 18) },
        });

        var result = Excel.Read<Customer>(path);
        Console.WriteLine($"  Imported {result.Rows.Count} customer(s); {result.Errors.Count} error(s).");
    }
}
