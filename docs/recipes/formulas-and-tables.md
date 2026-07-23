# Recipe: formulas inside a native Excel Table

Formula columns remain native formula cells when a schema is exported as a Table. CellSharp does not convert them to structured references or calculate their result.

```csharp
var schema = Excel.Schema<Order>()
    .SheetName("Orders")
    .AsTable("OrdersTable")
    .Column(x => x.Quantity)
    .Column(x => x.UnitPrice)
    .Column(x => x.Total, column => column
        .Formula(context => $"=A{context.Row}*B{context.Row}")
        .Format("#,##0.00"))
    .Build();

Excel.Write("orders.xlsx", orders, schema);
```

The `#,##0.00` format is invariant XLSX syntax and displays two decimal places using Excel's local separators, such as `1,00` in Italian Excel.

The first data row is Excel row 2. CellSharp requests recalculation only when it has actually written formula cells. It also stores the property's converted value as the formula's initial cache, allowing immediate import; Excel, LibreOffice, or another calculation-capable consumer replaces that cache on recalculation.

For an executable calculated-column example, see [Formulas and calculated columns](../use-cases/formulas-and-calculated-columns.md).
