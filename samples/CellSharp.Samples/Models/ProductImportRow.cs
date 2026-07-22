namespace CellSharp.Samples.Models;

internal sealed class ProductImportRow
{
    public string Name { get; set; } = string.Empty;

    public string? Category { get; set; }

    public int Quantity { get; set; }

    public DateTime AvailableFrom { get; set; }

    public string? Notes { get; set; }
}
