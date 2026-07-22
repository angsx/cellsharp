# Import and validate a user-edited workbook

## Scenario

Users have filled in the product-import template. Import the rows while preserving useful coordinates and messages for anything that needs correction.

## Model and schema

Use the `ProductImportRow` model and schema from [Data-entry workbook](data-entry-workbook.md). Its allowed values, numeric range, and date range run during import as well as in Excel.

## Read

```csharp
var result = Excel.Read(
    "uploaded-products.xlsx",
    schema,
    new ExcelReadOptions(ExcelInvalidRowPolicy.Include));

foreach (var product in result.Rows)
{
    Console.WriteLine(product.Name);
}

foreach (var error in result.Errors)
{
    Console.WriteLine($"{error.SheetName} {error.CellReference}: {error.Message}");
}
```

## Result

`Rows` contains converted products. `Errors` identifies conversion and validation issues with sheet, row, column, cell reference, header, and property information. `IsValid` is `true` only when `Errors` is empty.

`Skip` is the default policy and omits invalid rows. `Include` keeps partial invalid rows: successfully converted values remain, while failed conversions receive their CLR default. This is useful when an application needs to show a review screen instead of discarding a user’s work.

## Why this approach

Native Excel validation prevents many mistakes early, while CellSharp validates the uploaded file as the authoritative import boundary.

## Try it

Run [ImportRoundTrip](../../samples/CellSharp.Samples/Scenarios/ImportRoundTrip.cs). It writes deterministic product rows, reads them back, and prints the row/error counts. For every diagnostic field and policy, see [reading](../reading.md).
