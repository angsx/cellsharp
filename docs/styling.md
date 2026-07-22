# Styling

CellSharp applies `ExcelTheme.Default` to every exported workbook. Styling is deliberately restrained and never changes exported values or import conversion.

## Built-in themes

Choose a built-in theme through write options:

```csharp
Excel.Write("customers.xlsx", customers, options => options.Theme(ExcelTheme.Modern));
```

Available themes are `Default`, `Modern`, `Classic`, and `Minimal`. They define base fonts, header colors and weight, data backgrounds, alternate-row colors, alignment, and essential borders. The same options work with a schema:

```csharp
Excel.Write("customers.xlsx", customers, schema, options => options.Theme(ExcelTheme.Classic));
```

Internally, each built-in theme resolves to the same visual-template model used by custom styles.

## Custom reusable templates

`ExcelStyleTemplate` names a visual definition. The more generic `ExcelTemplate` is deliberately reserved for a future API that generates or consumes actual XLSX template files. `ExcelStyleTemplate` is sealed and immutable, so a `static readonly` instance is safe to reuse across independent exports.

```csharp
public static class CompanyExcelStyles
{
    public static readonly ExcelStyleTemplate Corporate = new(
        fontName: "Aptos",
        fontSize: 11,
        dataTextColor: "#202020",
        dataBackgroundColor: "#FFFFFF",
        headerTextColor: "#FFFFFF",
        headerBackgroundColor: "#1F4E78",
        headerBold: true,
        alternateRowBackgroundColor: "#F4F7FB",
        borderColor: "#B4C7E7");
}

Excel.Write("report.xlsx", data, options => options.Template(CompanyExcelStyles.Corporate));
```

The constructor exposes only the properties used today: data and header font name/size, data and header text/background colors, header boldness, alternate-row background, and border color. Colors use `#RRGGBB` and are validated when the template is created. `headerFontName` and `headerFontSize` default to the data font values.

`Theme(...)` is a convenient built-in template selector. If both `Theme(...)` and `Template(...)` are supplied, the custom template is always the base, regardless of call order.

## Header overrides and precedence

Override selected header values after choosing a theme or template. Unspecified values remain from the base.

```csharp
Excel.Write("customers.xlsx", customers, options => options
    .Template(CompanyExcelStyles.Corporate)
    .HeaderStyle(style => style
        .Bold()
        .Background("#202020")
        .Foreground("#FFFFFF")));
```

Colors use `#RRGGBB` and are validated before the workbook is created. The precedence is:

```text
Column format/alignment > explicit HeaderStyle values > custom template or built-in theme > library defaults
```

## Column formatting

Use the concise schema methods when presentation belongs to a known field:

```csharp
var schema = Excel.Schema<Order>()
    .Column(x => x.OrderDate, column => column
        .Header("Date")
        .Format("dd/MM/yyyy")
        .Width(14))
    .Column(x => x.Total, column => column
        .Header("Total")
        .Format("#,##0.00")
        .Align(ExcelHorizontalAlignment.Right))
    .Column(x => x.Discount, column => column.Format("0.00%"))
    .Build();
```

`Format` is an Excel format-code string. CellSharp stores it in the XLSX style table and deliberately does not parse it or use it for input conversion. Equivalent format codes are reused in the style catalog. `Width` must be greater than zero and no greater than 255. `Align` supports the small `General`, `Left`, `Center`, and `Right` enum.

## Widths and frozen headers

```csharp
Excel.Write("customers.xlsx", customers, options => options
    .AutoFitColumns()
    .FreezeHeaderRow());
```

`AutoFitColumns()` is a width estimate based on exported text, constrained to a practical minimum and maximum. It does not reproduce Excel application's exact autofit behavior. `FreezeHeaderRow()` freezes row one for easier scrolling. Explicit schema widths win over autofit:

```text
Width(...) > AutoFitColumns() estimation
```

## Alternating rows

```csharp
Excel.Write("customers.xlsx", customers, options => options
    .Template(CompanyExcelStyles.Corporate)
    .AlternatingRows());
```

The first data row uses the base data background and every second data row uses `AlternateRowBackgroundColor` from the active theme or template. If a template leaves that color null, CellSharp falls back to its data background instead of hard-coding a color in writer logic.

Generated XLSX templates contain no data rows, so `AlternatingRows()` has no visible effect there. Their column defaults still use the selected theme or visual template. See [generated templates](templates.md) for the empty-workbook behavior.

## Limits

Column format, width, and alignment are available, but arbitrary per-cell or row styling, conditional formatting, charts, and external XLSX templates are not available yet. For data structure and validation, see [schemas](schemas.md); for export behavior, see [writing](writing.md).
