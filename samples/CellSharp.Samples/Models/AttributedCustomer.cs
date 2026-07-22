using CellSharp;

namespace CellSharp.Samples.Models;

internal sealed class AttributedCustomer
{
    [ExcelColumn("Customer ID", Order = 1)]
    public int Id { get; set; }

    [ExcelColumn("Company", Order = 2, Optional = true)]
    public string? Name { get; set; }

    [ExcelColumn("Joined", Order = 3, Format = "yyyy-mm-dd")]
    public DateTime CreatedAt { get; set; }

    [ExcelIgnore]
    public string? InternalCode { get; set; }
}
