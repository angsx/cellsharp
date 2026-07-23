# Generated XLSX templates

`Excel.CreateTemplate` creates an empty XLSX workbook from an existing `ExcelSchema<T>`. It is intended for a user to download, complete manually in Excel or LibreOffice, and later import with the same schema.

## Basic usage

```csharp
Excel.CreateTemplate("customers-template.xlsx", schema);
```

The file contains one worksheet and a styled header row, but no data rows. Its name comes from `schema.SheetName`: `Sheet1` by default or the value configured with `.SheetName("Customers")`. Like `Excel.Write`, it deterministically replaces an existing file at the target path.

For a multi-sheet template, use `Excel.Workbook().AddTemplateSheet(schema)...CreateTemplate(path)`. It shares the same schemas and per-sheet presentation options as multi-sheet export. See [multi-sheet workbooks](multi-sheet.md).

`CreateTemplate` also accepts a writable, seekable stream and follows the same ownership and position rules as `Write`: it truncates from position zero, leaves the stream open, and finishes at the workbook end.

An immutable runtime schema overlay also applies to templates: `Excel.CreateTemplate(path, schema, overlay)` writes its effective headers and only its enabled columns. Native validation, formats, widths, and alignment are emitted only for those active columns.

## Styling

The existing write options apply without a separate template-specific options type:

```csharp
Excel.CreateTemplate(
    "customers-template.xlsx",
    schema,
    options => options
        .Theme(ExcelTheme.Modern)
        .FreezeHeaderRow());
```

An immutable reusable visual template works in the same way:

```csharp
Excel.CreateTemplate(
    "customers-template.xlsx",
    schema,
    options => options.Template(CompanyExcelStyles.Corporate));
```

`HeaderStyle(...)` overrides the selected theme or `ExcelStyleTemplate`; a custom `Template(...)` remains the base when both it and `Theme(...)` are specified. See [styling](styling.md) for complete precedence.

## What comes from the schema

The generated worksheet uses the schema's declared order and headers. `Optional()` columns are included: optional means an imported workbook may omit that header, not that the recommended blank template should omit it. `Ignore()` columns are excluded.

`Format(...)`, `Width(...)`, `Align(...)`, and `VerticalAlign(...)` are carried into the template. `Align(...)` and `VerticalAlign(...)` also determine the matching header alignment unless `HeaderAlign(...)` or `HeaderVerticalAlign(...)` overrides it. CellSharp stores a default XLSX style on each column rather than adding empty placeholder rows. This keeps the workbook empty while letting cells entered manually inherit column format and alignment in spreadsheet applications. `DateTime` columns receive CellSharp's usual date display format and `decimal` columns receive `0.00` when no explicit format is configured. Format codes use invariant XLSX syntax, so `0.00` displays as `1,00` for a value of `1` in Italian Excel.

## Write options on empty workbooks

`AutoFitColumns()` estimates width from headers only; future values cannot be predicted. Explicit schema `Width(...)` always wins. `FreezeHeaderRow()` freezes row one exactly as it does during export. `AlternatingRows()` has no visible effect because generated templates intentionally contain no data rows.

## Declarative validation

Declarative schema constraints become native Excel Data Validation on the full future column range, from row 2 through Excel's maximum worksheet row. They do not create placeholder rows:

```csharp
var schema = Excel.Schema<Customer>()
    .Column(x => x.Status, column => column
        .AllowedValues("Active", "Inactive", "Pending"))
    .Column(x => x.Age, column => column.Range(18, 120))
    .Build();
```

The status column becomes an Excel dropdown and the age column receives an inclusive numeric constraint. List values are stored in an internal hidden worksheet, avoiding fragile inline-list formula limits. `.Validate(...)` remains an arbitrary runtime .NET predicate and is not translated to Excel. See [validation](validation.md).

`ConvertWith(...)` does not create visible template content, but it remains compatible with the template's format, width, alignment, and independently compatible native validation. See [custom converters](converters.md).

## Limitations

Generated templates do not currently include formulas, charts, external XLSX templates, or editing of an existing workbook. Native validation is limited to declared lists, numeric ranges, and date ranges; custom predicates, cross-property rules, and conditional formatting are not translated.
