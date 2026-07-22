namespace CellSharp.Samples.Models;

internal sealed class OrderLine
{
    public string Product { get; set; } = string.Empty;

    public int Quantity { get; set; }

    public decimal UnitPrice { get; set; }

    public decimal Total { get; set; }
}
