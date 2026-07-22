# Export a list of objects

## Scenario

You have a list of customers and need a spreadsheet quickly. The property names are acceptable as headers and no special formatting is needed.

## Model

```csharp
public sealed class Customer
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public DateTime CreatedAt { get; set; }
}
```

## Write

```csharp
var customers = new[]
{
    new Customer { Id = 1, Name = "Acme Ltd", CreatedAt = new DateTime(2025, 1, 15) },
    new Customer { Id = 2, Name = "Contoso", CreatedAt = new DateTime(2025, 2, 8) },
};

Excel.Write("customers.xlsx", customers);
```

## Result

CellSharp creates `Sheet1` with `Id`, `Name`, and `CreatedAt` headers followed by the two rows. Public scalar properties are exported in reflection order; null values become blank cells.

## Why this approach

Conventions are the shortest path when the workbook is not yet a formal contract. Introduce a schema only when its shape needs to be controlled.

## Try it

Run the [BasicExport scenario](../../samples/CellSharp.Samples/Scenarios/BasicExport.cs). For conversion rules, see [writing](../writing.md).
