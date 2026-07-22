namespace CellSharp.Samples.Models;

internal sealed class Customer
{
    public int Id { get; set; }

    public string Name { get; set; } = string.Empty;

    public DateTime CreatedAt { get; set; }
}
