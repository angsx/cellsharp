# Worksheet settings

Schemas can carry a small set of native XLSX worksheet settings. They apply to writes, multi-sheet workbooks, and generated templates without requiring Excel to be installed.

```csharp
var schema = Excel.Schema<Order>()
    .SheetName("Order Details")
    .AutoFilter()
    .FreezePanes(rows: 1, columns: 2)
    .Landscape()
    .FitToPage(width: 1, height: 0)
    .RepeatHeaderRowOnPrint()
    .PrintGridlines()
    .Column(x => x.Product)
    .Column(x => x.Quantity)
    .Build();
```

## Filters and panes

`AutoFilter()` adds a worksheet-level native filter over exactly the generated header and data range. An empty export or template uses the header-only range, such as `A1:C1`. If the schema also uses `AsTable()`, CellSharp does not write a duplicate worksheet filter: the native Table AutoFilter is retained instead.

`FreezePanes(rows, columns)` freezes top rows and left columns. Both arguments must be zero or greater; `FreezePanes(0, 0)` writes no frozen pane. Existing `ExcelWriteOptionsBuilder.FreezeHeaderRow()` remains supported and freezes one row when the schema does not configure panes explicitly.

## Printing

`Landscape()` and `Portrait()` select orientation. `FitToPage(width, height)` enables native fit-to-page; a zero width or height means automatic sizing for that dimension, while negative values are rejected. `RepeatHeaderRowOnPrint()` writes the workbook-local `_xlnm.Print_Titles` defined name for row one, quoting worksheet names correctly. `PrintGridlines()` controls printed gridlines only and does not change the gridline setting shown in Excel's worksheet view.

Settings are independent per sheet in `Excel.Workbook()`. They preserve normal direct cell formatting, declarative validation, formulas, runtime overlays, and streams. Formula recalculation remains tied only to exported formula cells.

This is intentionally not a general print-layout API: CellSharp does not configure print areas, margins, paper sizes, zoom, or split panes.

# Headers, footers, comments, and page breaks

Worksheet utilities use native Excel structures and remain independent from schema settings.

```csharp
sheet.Header.Left("CellSharp");
sheet.Header.Center("Monthly sales");
sheet.Footer.Right("Page &P of &N");

sheet.Cell("C5").Comment("Imported from ERP", author: "CellSharp");
sheet.Row(50).PageBreakAfter();
sheet.Column("F").PageBreakAfter();
```

`Comment` produces an interoperable classic Excel note (not a threaded comment). Authors are deduplicated per worksheet; omitted authors use `CellSharp`. A second comment on the same cell fails fast. A comment can coexist with a style, hyperlink, or top-left merged cell.

Header/footer sections are `Left`, `Center`, and `Right`. Common Excel tokens are accepted: `&P` current page, `&N` total pages, `&D` date, `&T` time, `&F` filename, and `&A` sheet name. Ordinary ampersands are escaped for Excel.

`PageBreakAfter` uses absolute worksheet coordinates and is idempotent. It supports hidden rows/columns and XLSX limit coordinates. Headers, footers, and breaks coexist with orientation, fit-to-page, print gridlines, and repeated header rows.
