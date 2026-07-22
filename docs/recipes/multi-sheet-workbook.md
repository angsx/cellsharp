# Recipe: create a multi-sheet workbook

Compose typed sheets and write them through one workbook builder.

```csharp
Excel.Workbook()
    .AddSheet(customers, customerSchema)
    .AddSheet(orders, orderSchema)
    .Write("business-data.xlsx");
```

Each schema keeps its own worksheet name, Table settings, validation, formulas, and print/view configuration. To read the generated workbook once, use `Excel.Open(...)` and read each schema by name.

```csharp
using var workbook = Excel.Open("business-data.xlsx");
var importedCustomers = workbook.Read(customerSchema);
var importedOrders = workbook.Read(orderSchema);
```

Use `AddTemplateSheet(...)` only with `.CreateTemplate(...)`; a workbook write contains only data sheets.
