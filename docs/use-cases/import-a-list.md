# Import a list

## Scenario

You receive a simple workbook with headers matching your public property names. No custom headers, conversion, or validation rules are needed yet.

## Model

```csharp
public sealed class Customer
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public DateTime CreatedAt { get; set; }
}
```

## Read

```csharp
var result = Excel.Read<Customer>("customers.xlsx");

foreach (var customer in result.Rows)
{
    Console.WriteLine(customer.Name);
}

foreach (var error in result.Errors)
{
    Console.WriteLine($"{error.CellReference}: {error.Message}");
}
```

## Result

CellSharp matches headers to writable public properties without case sensitivity. `Rows` contains fully converted rows and `Errors` explains invalid cells. `IsValid` is true when there are no errors.

## Why this approach

This is convention-based import—the counterpart to zero-configuration export. Move to a schema when headers, worksheet selection, mapping, validation, or presentation become part of the file contract.

## Try it

Run [BasicImport](../../samples/CellSharp.Samples/Scenarios/BasicImport.cs). For invalid-row policies and localized text values, see [reading](../reading.md).
