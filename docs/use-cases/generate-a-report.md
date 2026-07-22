# Generate a report

## Scenario

Generate a small sales report that is immediately readable when opened, rather than a bare data dump.

## Model

```csharp
public sealed class ReportRow
{
    public string Customer { get; set; } = "";
    public int Orders { get; set; }
    public decimal Revenue { get; set; }
    public DateTime ReportedAt { get; set; }
}
```

## Schema and write

```csharp
var schema = Excel.Schema<ReportRow>()
    .SheetName("June report")
    .AsTable("JuneReport", "TableStyleMedium2")
    .FreezePanes(1, 0)
    .Landscape()
    .FitToPage(1, 0)
    .RepeatHeaderRowOnPrint()
    .Column(x => x.Customer, column => column.Width(28))
    .Column(x => x.Orders, column => column.Format("0").Width(10))
    .Column(x => x.Revenue, column => column.Format("#,##0.00").Width(16))
    .Column(x => x.ReportedAt, column => column.Header("As of").Format("yyyy-mm-dd"))
    .Build();

Excel.Write("sales-report.xlsx", report, schema, options => options.AlternatingRows());
```

## Result

The workbook has an Excel Table with filters, readable number/date formats, frozen headers, sized columns, alternating rows, and basic print settings.

## Why this approach

Schemas own durable worksheet structure; write options supply operation-level presentation. Tables are native Excel Tables, not a separate CellSharp data model.

## Try it

Run [Report](../../samples/CellSharp.Samples/Scenarios/Report.cs). See [Excel Tables](../tables.md) and [worksheet settings](../worksheet-settings.md).
