# Reusable domain schema

## Scenario

Your customer export is now shared with another team. It needs friendly headers, a stable column order, widths, and date formatting.

## Model

```csharp
public sealed class Customer
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public DateTime CreatedAt { get; set; }
}
```

## Schema

```csharp
var schema = Excel.Schema<Customer>()
    .SheetName("Customers")
    .Column(x => x.Id, column => column.Header("Customer ID").Width(14))
    .Column(x => x.Name, column => column.Header("Company").Width(30))
    .Column(x => x.CreatedAt, column => column
        .Header("Created")
        .Format("yyyy-mm-dd")
        .Width(14))
    .Build();
```

## Write / Read

```csharp
Excel.Write("customers.xlsx", customers, schema, options => options.FreezeHeaderRow());
var imported = Excel.Read<Customer>("customers.xlsx", schema);
```

## Result

The export has a `Customers` worksheet with the stated headers and presentation. The same schema reads those headers back, so write and import share one contract.

## Why this approach

Fluent configuration belongs close to the operation and can express dynamic behavior such as validators, converters, formulas, and runtime mappings.

## Try it

Run [FluentSchema](../../samples/CellSharp.Samples/Scenarios/FluentSchema.cs). See [typed schemas](../schemas.md) for all schema options.
