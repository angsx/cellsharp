namespace CellSharp.Samples.Models;

internal sealed class ReportRow
{
    public string Customer { get; set; } = string.Empty;

    public int Orders { get; set; }

    public decimal Revenue { get; set; }

    public DateTime ReportedAt { get; set; }
}
