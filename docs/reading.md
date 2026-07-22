# Reading XLSX files

`Excel.Read<T>` reads the first worksheet into an `ExcelReadResult<T>` when no schema is supplied:

```csharp
var result = Excel.Read<Customer>("customers.xlsx");
```

When a schema uses `.SheetName("Customers")`, `Read` selects that worksheet by name, case-insensitively. A missing configured sheet throws `InvalidOperationException`; CellSharp never silently falls back to the first sheet.

For multiple typed sheets in one package, open it once with `Excel.Open(path)` and call `workbook.Read(schema)` for each schema. `workbook.ReadAt(index, schema)` supports a zero-based public sheet index and verifies the selected name matches the schema. See [multi-sheet workbooks](multi-sheet.md).

`T` must have a public parameterless constructor and public writable non-indexer scalar properties. The supported types are listed in [compatibility](compatibility.md).

## Streams

All `Read` overloads also accept a `Stream`, with the same schema, overlay, and `ExcelReadOptions` combinations:

```csharp
var result = Excel.Read<Product>(uploadedStream);
var configured = Excel.Read(uploadedStream, schema, overlay, options);
```

The stream must be readable and seekable. CellSharp sets `Position` to zero before opening the XLSX package and never disposes a stream supplied by the caller. The stream position after a read is controlled by OpenXML package access and should not be used as an application contract; reset it explicitly before another consumer if needed.

## Untrusted input and resource limits

Reads enforce bounded package, Open XML part, physical-row, and collected-error defaults. `ExcelReadOptions` exposes `MaxPackageBytes`, `MaxCharactersInPart`, `MaxRows`, and `MaxErrors` for applications with a justified larger workload. See the [security model](security.md) before raising them.

For multi-sheet reads, pass limits when opening the package:

```csharp
var options = new ExcelReadOptions(maxRows: 250_000);
using var workbook = Excel.Open(uploadedStream, options);
var result = workbook.Read(customerSchema);
```

Package and Open XML part limits are fixed when `Excel.Open` opens the package. A later per-sheet options value controls row limits, error limits, and invalid-row handling for that read.

## Mapping and conversion

Without a schema, property names are required headers. Header matching is case-insensitive and column order does not matter. Extra columns and empty header cells are ignored. Duplicate non-empty headers and missing required headers are reported as structured errors and no rows are materialized.

With an explicit schema, CellSharp selects the first non-empty row that contains all required configured headers. A workbook generated with report content above `DataStartAt(...)` therefore round-trips through the same schema: the title and note are ignored and the offset schema header is used. A schema that maps every field exclusively with `MapFromColumn(...)` has no header identity to discover, so its first non-empty row remains the header by design.

By default, cells are converted using invariant conventions. Numbers do not accept group separators. Booleans accept `1`, `0`, `true`, and `false` without case sensitivity. `DateTime` accepts recognized Excel date cells and a limited set of ISO text formats. GUIDs are parsed without culture-specific behavior. A configured fallback culture can additionally parse text-form numbers and dates as described below.

CellSharp derives the used input from physical worksheet rows and cells, not from the worksheet dimension reference. Leading, trailing, and styled-only empty rows are ignored, even if an application declares a very large used range. A row is significant when it contains a non-empty value, a formula, or an Excel error cell. Whitespace is preserved as string data; a whitespace-only numeric, date, boolean, or GUID value is invalid.

Completely empty rows are ignored. Nullable value-type cells may be blank; blank non-nullable value-type cells produce an error. A physically absent string cell maps to `null`, while an explicit empty inline or shared string maps to `""` when another cell makes the row significant. Sparse cells retain their actual column positions.

Formula cells are read from their cached value through the normal conversion, converter, and validation pipeline. A formula without a cached value and an Excel error cell are reported as `InvalidValue`, while other valid rows continue to be imported. With `ExcelInvalidRowPolicy.Include`, the partial object is retained with the failing property left at its natural CLR default. This does not make CellSharp a formula engine; see [formulas](formulas.md).

## Localized text values

Native XLSX numbers and date serials are culture-independent. Some third-party systems instead export display values such as `45,498933` or `31/12/2026` as text. Configure `ExcelReadOptions.Culture` to accept those values:

```csharp
using System.Globalization;

var options = new ExcelReadOptions(
    culture: CultureInfo.GetCultureInfo("it-IT"),
    emptyStringAsNull: true);

var result = Excel.Read<Store>("external-export.xlsx", options);
```

`Culture` defaults to `CultureInfo.InvariantCulture` and is a fallback, not a replacement for the XLSX rules:

1. Numeric values are attempted with the invariant culture first, then with the configured culture. Group separators remain unsupported; the fallback is intended for localized decimal separators.
2. `DateTime` first accepts recognized Excel date serials and the supported invariant ISO text formats, then tries the configured culture.
3. String, Boolean, and GUID conversion is unchanged. The option does not localize headers, schema labels, validation messages, or exported values.

This precedence means a standards-compliant XLSX file behaves the same even when a fallback culture is supplied. The culture also applies to a custom converter's scalar `TCellValue` when that scalar is numeric or `DateTime`; CellSharp performs this scalar conversion before calling `TryRead`.

## Explicit empty strings as null

`EmptyStringAsNull` defaults to `false`. Set it to `true` when an external feed represents missing values as explicit zero-length inline or shared strings rather than absent cells. It runs before scalar and custom conversion:

| Incoming cell | Default (`false`) | `EmptyStringAsNull: true` |
| --- | --- | --- |
| Physically absent cell | `null` for `string` and nullable value types; `RequiredValueMissing` otherwise | Same |
| Explicit `""` to `string` | `""` | `null` |
| Explicit `""` to a nullable value type | `InvalidValue` | `null` |
| Explicit `""` to a non-nullable value type | `InvalidValue` | `RequiredValueMissing` |

The option matches only a zero-length string. Whitespace is not empty: it remains unchanged for a string target and is normally invalid for Boolean, GUID, and numeric targets. An empty cell does not make an otherwise empty worksheet row significant, so the setting matters when another cell (or a formula/error cell) causes that row to be imported.

Pass the options to any matching path or stream `Excel.Read(...)` overload. For multi-sheet imports, options supplied to `Excel.Open(...)` become the defaults for `workbook.Read(...)` and `ReadAt(...)`; a per-read options argument overrides them for that worksheet. Package and XML-part size limits, however, can only be enforced when the package is opened.

## Results and errors

```csharp
var result = Excel.Read<Customer>("customers.xlsx");

foreach (var customer in result.Rows)
{
    // Only rows with no conversion errors.
}

foreach (var error in result.Errors)
{
    Console.WriteLine(
        $"{error.SheetName} {error.CellReference} " +
        $"({error.Header}): {error.Message}");
}
```

`IsValid` is true only when `Errors` is empty. `ExcelReadError` provides the sheet name, one-based visible `Row`, optional one-based physical `Column` and cell reference, header, CLR `PropertyName`, raw value, expected .NET type, a typed `Code`, `Kind`, and a readable message. `Kind` distinguishes missing/duplicate headers, required values, conversion, and validation; `Code` remains the detailed compatible code (`MissingHeader`, `DuplicateHeader`, `RequiredValueMissing`, `InvalidValue`, or `ValidationFailed`). Header-level errors identify row 1, the visible header row, and have no physical column when the required header is absent.

Rows with conversion errors are excluded from `Rows`; valid rows from the same workbook remain available. Missing files, invalid XLSX packages, unavailable worksheets, and unsupported target types are usage or structural failures and throw exceptions.

## Invalid-row policy

The historical default is `ExcelInvalidRowPolicy.Skip`: a row with any conversion or validation error is not materialized. Header-shape errors still prevent all row materialization because there is no reliable mapping.

Use a small immutable options value to retain invalid data rows for review:

```csharp
var result = Excel.Read(
    "products.xlsx",
    schema,
    new ExcelReadOptions(ExcelInvalidRowPolicy.Include));
```

`Include` returns the partially populated CLR object: values that converted successfully are assigned, while a failed conversion leaves that property at its natural CLR default (`null` or `default(TProperty)`). Validation failures retain their converted value. Errors are always collected and `IsValid` remains false. The same options are available on `Excel.Open(...).Read(...)` and `ReadAt(...)`.

## Schema validation

A schema column can validate a value after CellSharp has converted it successfully:

```csharp
var schema = Excel.Schema<Customer>()
    .Column(x => x.Age, column => column
        .Validate(age => age >= 18, "Customer must be at least 18"))
    .Build();
```

An unparseable value produces `InvalidValue`; a converted value that fails a predicate produces `ValidationFailed`. Both exclude the row, while other valid rows remain in `Rows`. Multiple rules on a column are all evaluated and each failure is reported. Nullable properties pass their converted `null` value to the predicate.

Validator exceptions are not converted to data errors: they propagate to the caller because they indicate an application or configuration problem. Use a [schema](schemas.md) to customize headers, choose columns, make a declared column optional, or add validation.

Declarative schema constraints run before custom predicates and return the same `ValidationFailed` result when violated. They can also be emitted as native Excel Data Validation. See [validation](validation.md).

When a schema column uses `ConvertWith`, CellSharp converts the raw cell to the converter's declared scalar type and then calls `TryRead`. A `false` result becomes a `Conversion` diagnostic (`InvalidValue` code); converter exceptions propagate because they indicate application bugs or impossible converter state rather than expected workbook data. See [custom converters](converters.md).

## Runtime mappings

Schema columns can select their incoming data at runtime while retaining the same converters and validation rules:

```csharp
var schema = Excel.Schema<Product>()
    .Column(x => x.Color, column => column.MapFromColumn(3))
    .Column(x => x.Size, column => column.MapFromHeader("size"))
    .Build();
```

`MapFromColumn` uses **one-based Excel column numbers** (`1` is column A). `MapFromHeader` matches case-insensitively. Both settings affect import only; `Header(...)` still controls the generated export header.

When both are configured for a property, the positional mapping wins, then an active overlay header, then `MapFromHeader`, then its normal schema header. A positional mapping also claims that physical column before header mapping runs. Duplicate runtime column numbers, duplicate runtime headers, and a runtime header that collides with another property's schema header are rejected while building the schema. A required missing runtime mapping uses the existing `MissingHeader` diagnostic; `Optional()` suppresses it.

## Per-operation schema overlay

Pass an `ExcelSchemaOverlay<T>` when an operation expects a final application-supplied header or enables only selected columns:

```csharp
var result = Excel.Read("products.xlsx", schema, overlay);
```

Import precedence is `MapFromColumn`, runtime overlay header, `MapFromHeader`, then schema `Header`. Disabled columns are absent from mapping entirely, so their physical presence or absence, converters, and validators are ignored. A `MapFromColumn(3)` always means physical Excel column C even if an earlier schema column is disabled.
