# Import diagnostics

`Excel.Read<T>` returns `ExcelReadResult<T>` rather than silently discarding workbook problems. `Rows` and `Errors` are immutable snapshots; `IsValid` is true only when `Errors` is empty.

```csharp
var result = Excel.Read("orders.xlsx", schema);

foreach (var error in result.Errors)
{
    Console.WriteLine(
        $"{error.SheetName} {error.CellReference} " +
        $"{error.PropertyName} [{error.Kind}/{error.Code}]: {error.Message}");
}
```

Coordinates exposed through `Row`, `Column`, `RowNumber`, and `ColumnNumber` are one-based Excel coordinates. `RawValue` is the original textual cell value used by the conversion pipeline; `ExpectedType` identifies the CLR target when conversion applies. Treat `RawValue`, `Header`, and `SheetName` as untrusted input and encode them for the destination before rendering or logging them.

The default `ExcelInvalidRowPolicy.Skip` omits an invalid data row. `ExcelInvalidRowPolicy.Include` keeps its partially populated object: successfully converted properties remain assigned, while a conversion failure leaves the affected property at its CLR default. Errors are collected in both modes.

`Kind` is the stable broad category (`MissingHeader`, `DuplicateHeader`, `RequiredValue`, `Conversion`, or `Validation`); `Code` retains the detailed compatible error code. Structural failures such as a missing file, an invalid XLSX package, or an unavailable worksheet are usage/package errors and throw exceptions instead of being represented as row diagnostics.
