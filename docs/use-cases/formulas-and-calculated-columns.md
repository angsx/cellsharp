# Formulas and calculated columns

## Scenario

An order report needs `Total = Quantity × Unit price` visible and recalculable in Excel.

## Model

```csharp
public sealed class OrderLine
{
    public string Product { get; set; } = "";
    public int Quantity { get; set; }
    public decimal UnitPrice { get; set; }
    public decimal Total { get; set; }
}
```

## Schema and write

```csharp
var schema = Excel.Schema<OrderLine>()
    .SheetName("Order report")
    .Column(x => x.Product)
    .Column(x => x.Quantity)
    .Column(x => x.UnitPrice, column => column.Format("#,##0.00"))
    .Column(x => x.Total, column => column
        .Formula(context => $"=B{context.Row}*C{context.Row}")
        .Format("#,##0.00"))
    .Build();

Excel.Write("order-report.xlsx", lines, schema);
```

## Result

`Total` is a native Excel formula cell. CellSharp writes the property value as its initial cached result and requests recalculation when formula cells are exported. `#,##0.00` is invariant XLSX syntax, so Excel renders whole totals as `1,00` in an Italian locale.

## Why this approach

CellSharp writes formulas; it does not calculate them. Excel, LibreOffice, or another spreadsheet engine recalculates on open. Keep the model’s `Total` value consistent with the expression for consumers that read before recalculation.

## Try it

Run [FormulaColumns](../../samples/CellSharp.Samples/Scenarios/FormulaColumns.cs). See [formulas](../formulas.md) for import cache behavior and formula safety.
