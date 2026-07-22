namespace CellSharp.Samples.Models;

internal sealed class CustomerWithStatus
{
    public int Id { get; set; }

    public string Name { get; set; } = string.Empty;

    public CustomerStatus? Status { get; set; }
}
