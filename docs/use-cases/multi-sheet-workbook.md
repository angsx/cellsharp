# Multi-sheet workbook

## Scenario

Finance needs one workbook containing a customer list and its order lines. Each worksheet has a different model and independently controlled presentation.

## Schemas

```csharp
var customerSchema = Excel.Schema<Customer>()
    .SheetName("Customers")
    .Column(x => x.Id, column => column.Header("Customer ID"))
    .Column(x => x.Name, column => column.Header("Company"))
    .Build();

var orderSchema = Excel.Schema<OrderLine>()
    .SheetName("Orders")
    .AsTable("OrdersTable")
    .FreezePanes(1, 0)
    .Column(x => x.Product)
    .Column(x => x.Quantity)
    .Column(x => x.UnitPrice, column => column.Format("#,##0.00"))
    .Build();
```

The `#,##0.00` code is written in invariant XLSX syntax. Excel applies the locale when displaying it, so it renders `1` as `1,00` in Italian.

## Write / Read

```csharp
Excel.Workbook()
    .AddSheet(customers, customerSchema, options => options.FreezeHeaderRow())
    .AddSheet(orders, orderSchema, options => options.AlternatingRows())
    .Write("customer-orders.xlsx");

using var workbook = Excel.Open("customer-orders.xlsx");
var importedCustomers = workbook.Read(customerSchema);
var importedOrders = workbook.Read(orderSchema);
```

## Result

The output contains `Customers` and `Orders`; the second sheet is an Excel Table. `Excel.Open` opens the package once and each schema selects its sheet by name.

## Why this approach

Use the single-sheet API by default. The workbook builder becomes useful when a single XLSX package truly has several typed document shapes.

## Try it

Run [MultiSheetWorkbook](../../samples/CellSharp.Samples/Scenarios/MultiSheetWorkbook.cs). See [multi-sheet workbooks](../multi-sheet.md) for template sheets and stream overloads.
