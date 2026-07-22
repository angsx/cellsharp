# Data-entry workbook

## Scenario

Operations staff need a blank product-import file that guides entry before it returns to the application. The workbook should have useful labels, dropdowns, ranges, date limits, formatting, widths, and a frozen header.

## Model

```csharp
public sealed class ProductImportRow
{
    public string Name { get; set; } = "";
    public string? Category { get; set; }
    public int Quantity { get; set; }
    public DateTime AvailableFrom { get; set; }
    public string? Notes { get; set; }
}
```

## Schema

```csharp
var schema = Excel.Schema<ProductImportRow>()
    .SheetName("Product import")
    .FreezePanes(1, 0)
    .Column(x => x.Name, column => column.Header("Product name").Width(30))
    .Column(x => x.Category, column => column
        .Header("Category").Optional()
        .AllowedValues("Hardware", "Software", "Service").Width(16))
    .Column(x => x.Quantity, column => column
        .Header("Opening quantity").Range(0, 100000).Format("0").Width(18))
    .Column(x => x.AvailableFrom, column => column
        .Header("Available from")
        .DateBetween(new DateTime(2025, 1, 1), new DateTime(2030, 12, 31))
        .Format("yyyy-mm-dd").Width(16))
    .Column(x => x.Notes, column => column.Header("Notes").Optional().Width(36))
    .Build();
```

## Write

```csharp
Excel.CreateTemplate("product-import-template.xlsx", schema, options => options
    .Theme(ExcelTheme.Modern)
    .FreezeHeaderRow());
```

## Result

The header-only workbook embeds native Excel dropdown, numeric, and date validation for future data rows. `Optional()` means the header may be absent when importing a third-party sheet; it does not make a value required or optional in Excel.

## Why this approach

The same schema later validates imported user edits, keeping the spreadsheet contract in one place.

## Try it

Run [DataEntryTemplate](../../samples/CellSharp.Samples/Scenarios/DataEntryTemplate.cs), then see [validation](../validation.md) and [templates](../templates.md).
