using CellSharp;

namespace CellSharp.Samples.Models;

internal sealed class CustomerStatus
{
    internal static CustomerStatus Active { get; } = new("ACTIVE");

    internal static CustomerStatus Pending { get; } = new("PENDING");

    private CustomerStatus(string code)
    {
        Code = code;
    }

    internal string Code { get; }

    internal static CustomerStatus? FromCode(string code) => code switch
    {
        "ACTIVE" => Active,
        "PENDING" => Pending,
        _ => null,
    };
}

internal sealed class CustomerStatusConverter : IExcelValueConverter<CustomerStatus?, string>
{
    internal static CustomerStatusConverter Instance { get; } = new();

    public string Write(CustomerStatus? value) => value?.Code ?? string.Empty;

    public bool TryRead(string value, out CustomerStatus? converted)
    {
        converted = CustomerStatus.FromCode(value);
        return converted is not null;
    }
}
