# Attribute-based schema

## Scenario

A DTO has one stable Excel representation that several application paths use. Put the durable mapping metadata on the type and opt in to it explicitly.

## Model

```csharp
public sealed class Customer
{
    [ExcelColumn("Customer ID", Order = 1)]
    public int Id { get; set; }

    [ExcelColumn("Company", Order = 2, Optional = true)]
    public string? Name { get; set; }

    [ExcelIgnore]
    public string? InternalCode { get; set; }
}
```

## Schema and write

```csharp
var schema = Excel.SchemaFromAttributes<Customer>(builder => builder
    .SheetName("Customers")
    .Column(x => x.Name, column => column.Width(30)));

Excel.Write("customers.xlsx", customers, schema);
```

## Result

The workbook contains `Customer ID` and `Company`, in that order, and excludes `InternalCode`. Attributes are opt-in: schema-less `Excel.Write(path, rows)` and `Excel.Read<T>(path)` still use property conventions.

## Why this approach

Attributes suit stable, reusable mappings. The fluent overload can replace an attributed property’s defaults, as above; prefer fluent schemas for local or dynamic behavior.

## Try it

Run [AttributeSchema](../../samples/CellSharp.Samples/Scenarios/AttributeSchema.cs). The complete attribute rules are in [writing](../writing.md#attribute-based-schemas).
