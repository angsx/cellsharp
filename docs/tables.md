# Excel Tables

`AsTable(...)` turns a schema export into a native Excel Table. It is opt-in: schemas without it produce the same ordinary worksheets as before.

```csharp
var schema = Excel.Schema<Order>()
    .SheetName("Orders")
    .AsTable("OrdersTable")
    .Column(x => x.Product)
    .Column(x => x.Quantity)
    .Column(x => x.UnitPrice)
    .Column(x => x.Total, column => column
        .Formula(context => $"=B{context.Row}*C{context.Row}")
        .Format("#,##0.00"))
    .Build();

Excel.Write("orders.xlsx", orders, schema);
```

`#,##0.00` uses the invariant XLSX decimal placeholder. Excel renders it with the local separator, so an Italian installation shows `1,00` for a value of `1`.

## Names and styles

`AsTable()` generates a deterministic name from the worksheet name, such as `OrdersTable`. `AsTable("OrdersTable")` preserves an explicit name. Names are validated as Excel identifiers and are unique case-insensitively across a multi-sheet workbook; an explicit duplicate fails the write.

The optional second argument selects a built-in table style:

```csharp
.AsTable("OrdersTable", "TableStyleLight9")
```

The default is `TableStyleMedium2` with row stripes. Pass `style: null` to omit `TableStyleInfo`. CellSharp writes its direct header, cell, format, width, and alignment styles as usual; Excel applies its normal precedence rules when displaying them alongside the native table style.

## Range, filters, and empty data

The table range is exactly the active schema columns, from the header through the last exported row. Ignored or runtime-disabled columns are absent. Every table includes its native AutoFilter on the header.

An empty data export and `CreateTemplate` both create a valid header-only table such as `A1:C1`. CellSharp does not add dummy rows or formula cells to templates.

## Formulas, validation, and multi-sheet workbooks

Formula columns remain ordinary native formula cells inside a table; CellSharp does not rewrite them to structured references. A callback may deliberately return a structured-reference expression, but CellSharp does not parse it. Formula recalculation flags are still set only when actual formula cells are exported.

Declarative validation, including the internal `_CellSharpValidation` worksheet when needed, remains independent from table definitions. Tables work for any subset of sheets in `Excel.Workbook()`; each table has a workbook-global deterministic ID and name, while its definition part remains attached to its own worksheet.

There is no table import model, totals-row API, append/resize support, slicers, custom table styles, or structured-reference builder. `Excel.Read(...)` simply reads the worksheet cells as before, including files that also contain a table definition.

For a report that combines a Table with widths, frozen headers, and print settings, see [Generate a report](use-cases/generate-a-report.md).
