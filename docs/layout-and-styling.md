# Layout and styling

CellSharp keeps schema mapping and worksheet layout separate. Existing `Excel.Write` calls remain the shortest path for ordinary exports; use `Excel.Workbook()` when a worksheet also needs report-like content.

```csharp
Excel.Workbook()
    .AddSheet("Report", sheet =>
    {
        sheet.Merge("A1:D1")
            .Value("Sales report")
            .Style(style => style
                .Bold()
                .FontSize(18)
                .FontColor("#FFFFFF")
                .FillColor("#1F4E78")
                .AlignCenter()
                .VerticalAlignCenter());

        sheet.Range("A3:A8")
            .Merge()
            .Value("Europe")
            .Style(style => style.VerticalText().AlignCenter().VerticalAlignCenter());

        sheet.Cell("D10").Formula("=SUM(D4:D9)").NumberFormat("€ #,##0.00");
        sheet.Cell("A12").Value("OpenAI").Hyperlink("https://openai.com");
        sheet.Row(1).Height(30);
        sheet.Columns("A:D").Width(16);
    })
    .Write("report.xlsx");
```

Typed data and layout share the same writer. `DataStartAt` moves the header and data region while keeping formula contexts, tables, and native validations aligned with that region.

```csharp
Excel.Workbook()
    .AddSheet(customers, customerSchema, sheet =>
    {
        sheet.Merge("A1:F1").Value("Customers").Style(s => s.Bold().FontSize(18));
        sheet.DataStartAt("A3");
    })
    .Write("customers.xlsx");
```

Styles are reusable and deduplicated in the workbook style catalog.

```csharp
var title = ExcelStyle.Create(s => s.Bold().FontSize(18).AlignCenter());
sheet.Range("A1:D1").Style(title);
```

`NumberFormat(...)` uses the same invariant XLSX codes as schema `Format(...)`: use `.` for decimal placeholders. For example, `€ #,##0.00` displays a whole value as `€ 1,00` in Italian Excel.

## Style composition and precedence

Styles are sparse: a builder only overrides the properties it configures. The resolved order is workbook/theme, schema property, column, row, range, then cell. The last level that sets a property wins; properties it does not set continue to inherit. Explicit `false` is an override, so `Bold(false)` and `WrapText(false)` reliably clear an earlier setting. Borders compose independently on each side.

```csharp
Excel.Workbook().AddSheet(rows, schema, sheet =>
{
    sheet.Column("B").Style(s => s.FontSize(11));
    sheet.Range("A2:B100").Style(s => s.Bold());
    sheet.Cell("B2").Style(s => s.Bold(false).FontColor("#FF0000"));
});
```

In this example, `B2` keeps its schema number format and alignment, has an 11-point font, is not bold, and is red. Equivalent resolved styles share a single XLSX cell format.

## Selective AutoFit

AutoFit measures only materialized cells, so it is safe on sparse and layout-only sheets. Explicit widths always win over selective AutoFit.

```csharp
sheet.Columns("A:D").AutoFit();
sheet.Column("C").Width(25);
sheet.Range("F:H").AutoFitColumns();
```

The final precedence is explicit width, selective AutoFit, global `AutoFitColumns` export option, then Excel's default width. `DataStartAt` does not change the coordinate system used by selective AutoFit.

Coordinates are Excel one-based coordinates. Local A1 references only are accepted (`A1`, `A1:D8`); invalid references, inverted ranges, and overlapping merges fail immediately. A duplicate merge is harmless. Values and formulas in a range are applied to its cells; for a merged region place content in its top-left cell.

The current layer writes native XLSX cells, merge cells, composable styles, number formats, row/column metadata, selective AutoFit, hyperlinks, and [native conditional formatting](conditional-formatting.md). The remaining Excel-native capabilities are planned in the [toolkit milestone plan](next-toolkit-milestone.md).

Ranges can also declare [workbook-scoped named ranges](named-ranges.md), while cells support classic Excel comments. Printable headers, footers, manual page breaks, and [native images](images.md) are configured from the worksheet builder.

For recurring report structure, use the lightweight [report components](report-layouts.md). They expand into these same ranges, values, merges, and styles, so they remain fully composable with the layout API.
