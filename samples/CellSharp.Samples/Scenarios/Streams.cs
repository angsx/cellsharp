using CellSharp;
using CellSharp.Samples.Models;

namespace CellSharp.Samples.Scenarios;

internal static class Streams
{
    internal static void Run(string outputDirectory)
    {
        var customers = new[]
        {
            new Customer { Id = 301, Name = "Web API customer", CreatedAt = new DateTime(2025, 7, 1) },
        };
        var path = Path.Combine(outputDirectory, "stream-export.xlsx");

        using var stream = new MemoryStream();
        Excel.Write(stream, customers);
        File.WriteAllBytes(path, stream.ToArray());

        stream.Position = 0;
        var result = Excel.Read<Customer>(stream);
        Console.WriteLine($"  Stream round trip read {result.Rows.Count} customer(s).");
    }
}
