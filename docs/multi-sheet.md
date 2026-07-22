# Multi-sheet workbooks

The single-sheet API remains the shortest path:

```csharp
Excel.Write("customers.xlsx", customers, customerSchema);
var result = Excel.Read<Customer>("customers.xlsx", customerSchema);
```

Use `Excel.Workbook()` only when one XLSX file contains multiple typed worksheets.

## Write a workbook

Each worksheet has its own row type, schema, sheet name, and presentation options. Declaration order is workbook order.

```csharp
Excel.Workbook()
    .AddSheet(customers, customerSchema, options => options
        .Theme(ExcelTheme.Modern)
        .FreezeHeaderRow())
    .AddSheet(orders, orderSchema, options => options
        .AutoFitColumns())
    .Write("business-data.xlsx");
```

Every schema supplies its own `SheetName(...)`. Names must be valid Excel names and unique case-insensitively within the builder. The existing column converter, validation, formatting, `Optional`, and `Ignore` behavior is unchanged per sheet.

Each sheet may also receive its own immutable runtime overlay before its optional write configuration:

```csharp
Excel.Workbook()
    .AddSheet(customers, customerSchema, customerOverlay)
    .AddSheet(orders, orderSchema, orderOverlay)
    .Write("business-data.xlsx");
```

The overlays are type-safe and independent: a runtime header or disabled column on `Customers` cannot affect `Orders`.

`Write(stream)` and `CreateTemplate(stream)` are available for writable seekable streams. They use the same workbook pipeline, including the shared hidden validation lookup worksheet, and leave the caller stream open at the end.

Native list validations across the workbook use one hidden CellSharp lookup worksheet where possible. Equal value lists are deduplicated. The hidden worksheet is never included by `ReadAt` and cannot be read as application data.

## Read a workbook

Open once, then read each typed worksheet using its schema:

```csharp
using var workbook = Excel.Open("business-data.xlsx");

var customers = workbook.Read(customerSchema);
var orders = workbook.Read(orderSchema);
```

Read options passed while opening the package become the default for every sheet. This is useful when all sheets come from the same localized external system:

```csharp
using System.Globalization;

using var workbook = Excel.Open(
    "business-data.xlsx",
    new ExcelReadOptions(
        culture: CultureInfo.GetCultureInfo("it-IT"),
        emptyStringAsNull: true));

var customers = workbook.Read(customerSchema);
var orders = workbook.Read(orderSchema);
```

Pass an `ExcelReadOptions` value to an individual `Read` or `ReadAt` call to override the parsing, empty-string, row, error, and invalid-row settings for that sheet. Package and XML-part limits are fixed by the options used by `Excel.Open(...)`, because they are checked while opening the package.

Use `workbook.Read(schema, overlay)` (or the corresponding `ReadAt` overload) when the incoming headers and active columns are operation-specific.

`Read` selects the schema's `SheetName` case-insensitively, including an explicitly named hidden user worksheet. A missing name throws `InvalidOperationException`; there is no fallback to another worksheet.

`Excel.Open(stream)` opens a readable seekable stream once for multiple reads. Its `ExcelWorkbookReader` owns only the OpenXML package: disposing it releases package resources but leaves the caller's stream open.

`ReadAt` is available when the workbook position is the required contract:

```csharp
var customers = workbook.ReadAt(0, customerSchema);
```

Indexes are zero-based and count public worksheets only, excluding CellSharp's internal hidden validation worksheet. `ReadAt` verifies that the selected worksheet name equals `schema.SheetName`; it rejects mismatched schemas instead of silently reading the wrong data.

## Create a multi-sheet template

Use the same workbook builder and schemas for header-only templates:

```csharp
Excel.Workbook()
    .AddTemplateSheet(customerSchema, options => options.Theme(ExcelTheme.Modern))
    .AddTemplateSheet(orderSchema, options => options.AutoFitColumns())
    .CreateTemplate("business-template.xlsx");
```

Template worksheets retain their own headers, formatting, widths, alignment, frozen-header options, and native validation. `AddSheet(...)` is for `.Write(...)`; `AddTemplateSheet(...)` is for `.CreateTemplate(...)`, so a builder cannot accidentally mix the two operation kinds.

`AddTemplateSheet(schema, overlay)` applies the same runtime headers and conditional participation to a single template worksheet.

## Scope

The API intentionally has no public OpenXML types, formula engine, general workbook layout engine, chart, pivot, or arbitrary worksheet-mutation model. The workbook reader owns the opened package and is disposable, leaving future APIs additive.

For a complete customers-and-orders example, see [Multi-sheet workbook](use-cases/multi-sheet-workbook.md).
