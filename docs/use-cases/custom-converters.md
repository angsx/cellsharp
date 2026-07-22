# Custom converters

## Scenario

Your model uses a domain value instead of a CellSharp scalar. For example, a customer status is an application type but the spreadsheet should store a stable text code.

## Model and converter

```csharp
public sealed class Customer
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public CustomerStatus? Status { get; set; }
}

public sealed class CustomerStatus
{
    public string Code { get; init; } = "";
}

public sealed class CustomerStatusConverter
    : IExcelValueConverter<CustomerStatus?, string>
{
    public string Write(CustomerStatus? value) => value?.Code ?? string.Empty;

    public bool TryRead(string value, out CustomerStatus? converted)
    {
        converted = value switch
        {
            "ACTIVE" or "PENDING" => new CustomerStatus { Code = value },
            _ => null,
        };
        return converted is not null;
    }
}
```

## Schema, write, and read

```csharp
var schema = Excel.Schema<Customer>()
    .Column(x => x.Id)
    .Column(x => x.Name)
    .Column(x => x.Status, column => column.ConvertWith(new CustomerStatusConverter()))
    .Build();

Excel.Write("customers.xlsx", customers, schema);
var imported = Excel.Read<Customer>("customers.xlsx", schema);
```

## Result

The XLSX cell contains a supported scalar (`string`), while application code works with `CustomerStatus`. On import, an unknown status code becomes an `InvalidValue` diagnostic; a converter exception is treated as an application bug and propagates.

## Why this approach

Use a converter for reusable domain-to-cell translation. It works in both directions and keeps OpenXML details out of your model.

## Try it

Run [CustomConverter](../../samples/CellSharp.Samples/Scenarios/CustomConverter.cs). See [custom converters](../converters.md) for nullability, formats, and validation ordering.
