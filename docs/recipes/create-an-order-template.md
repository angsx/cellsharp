# Recipe: create an order template

Build one schema for both a data-entry template and later typed import.

```csharp
var schema = Excel.Schema<Order>()
    .SheetName("Orders")
    .AsTable("OrdersTable")
    .AutoFilter()
    .FreezePanes(1, 0)
    .Column(x => x.Customer)
    .Column(x => x.Status, column => column.AllowedValues("New", "Paid"))
    .Column(x => x.Quantity, column => column.Range(1, 1000))
    .Column(x => x.DeliveryDate, column => column.DateBetween(DateTime.Today, DateTime.Today.AddYears(1)))
    .Build();

Excel.CreateTemplate("orders-template.xlsx", schema);
```

The result is header-only. It contains native Table/validation metadata but no dummy data rows or prefilled formula cells.

For the end-to-end user workflow, see [Data-entry workbook](../use-cases/data-entry-workbook.md).
