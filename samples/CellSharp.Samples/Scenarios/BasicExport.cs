using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class BasicExport
{
    internal static void Run(string outputDirectory)
    {
        var customers = new[]
        {
            new Customer { Id = 1, Name = "Acme Ltd", CreatedAt = new DateTime(2025, 1, 15) },
            new Customer { Id = 2, Name = "Contoso", CreatedAt = new DateTime(2025, 2, 8) },
        };

        Excel.Write(Path.Combine(outputDirectory, "basic-export.xlsx"), customers);
    }
}
