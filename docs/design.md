# Design notes

## Scope

The first slices write an `IEnumerable<T>` to one XLSX worksheet and read the first worksheet into typed rows. Without a schema, that is the first worksheet; a schema may select one worksheet by name, case-insensitively. `Excel.Workbook()` generalizes the same typed schema model to multiple ordered worksheets, while `Excel.Open(...)` keeps one package open for multiple typed reads. Export headers come from public readable scalar properties; import maps headers to public writable scalar properties.

## Public boundary

`CellSharp.Excel` is the current public entry point. Open XML types remain internal implementation details, so the API does not commit consumers to a specific XLSX library. There is no backend abstraction yet: one implementation is simpler and testable, and an interface would not serve a second implementation today.

## Data representation

Text values, including text beginning with `=`, are stored as inline strings. Only a future explicit formula API may create formula cells. Numeric values and booleans use their native XLSX representations. `DateTime` is written as an OA date number with a small workbook style so spreadsheet applications recognize it as a date.

On import, numbers first use invariant XLSX-style parsing without group separators. Dates first use recognized Excel date cells or a small set of invariant ISO formats. `ExcelReadOptions.Culture` optionally adds an explicit, deterministic fallback for numeric and date values stored as text; the process never depends on the machine's current culture. Booleans accept `1`, `0`, `true`, and `false` case-insensitively, while GUID and string behavior is culture-independent. `ExcelReadOptions.EmptyStringAsNull` optionally normalizes an explicit zero-length string before scalar conversion; it does not treat whitespace as empty.

## Read result and diagnostics

`Excel.Read<T>` returns `ExcelReadResult<T>`. Its default `ExcelInvalidRowPolicy.Skip` preserves the historical behavior: `Rows` includes only rows whose mapped values all converted successfully. `ExcelInvalidRowPolicy.Include`, supplied through immutable `ExcelReadOptions`, returns a partially populated object for an invalid data row; failed conversions remain at CLR defaults while validation failures keep their converted value. `Errors` contains `ExcelReadError` values with sheet, one-based Excel coordinates, header, CLR property, raw value, expected type, detailed code, category, and message. `IsValid` is true when no errors were recorded regardless of row policy.

The current error codes are limited to `MissingHeader`, `DuplicateHeader`, `RequiredValueMissing`, `InvalidValue`, and `ValidationFailed`. Missing or duplicate headers prevent any data rows from being materialized. Invalid values affect only their row, allowing other rows to be returned.

File access, invalid XLSX packages, unavailable worksheets, and an unsupported target type are usage or structural failures and throw exceptions. Cell and row data issues are results, not exceptions.

## Header mapping

Header matching is case-insensitive and independent of column order. Without an explicit schema, all public writable supported properties are required headers. Unknown columns and empty or whitespace-only header cells are ignored. Any repeated non-empty header produces `DuplicateHeader` and stops row materialization, rather than choosing one duplicate arbitrarily. CellSharp does not trust worksheet dimension metadata: it scans actual row and cell elements, skips styled-only rows, and considers a data row significant only when it contains a non-empty value, a formula, or an Excel error cell. Partially populated rows are converted and can produce errors.

## Typed schemas

`Excel.Schema<T>()` creates a builder; `.Build()` returns an `ExcelSchema<T>` whose column definitions are copied into an immutable read-only representation. A built schema contains no operation state and can safely be reused for concurrent reads and writes.

Schemas are inclusive. When one is supplied, only properties declared with `.Column(...)` and not marked `.Ignore()` participate in the document. Declaration order determines export order. This prevents a newly added property on `T` from silently changing an established XLSX contract. A schema column must select one direct, public readable and writable scalar property; duplicate properties, duplicate effective headers, unsupported properties, and invalid expressions fail while building the schema.

`SheetName(...)` controls the one worksheet used by the schema across writing, reading, and template generation. It defaults to `Sheet1`; configured schemas select a matching name case-insensitively on read and fail explicitly if it is absent. `Header(...)` controls both the generated header and import mapping. Declared schema columns are required by default; `Optional()` permits an entire header to be absent and leaves the property at its .NET default. Column presence deliberately does not make nullable cells non-null: nullability continues to determine whether an empty cell can be assigned. This leaves future value-validation naming available for a distinct concern.

`MapFromColumn(int)` and `MapFromHeader(string)` are schema-column import overrides. `ExcelSchemaOverlay<T>` is the immutable per-operation layer for effective headers and column participation. Its precedence is positional mapping, overlay header, `MapFromHeader`, then schema header. Disabled columns are removed before both export and import metadata are built, so converter and validation work is not invoked; physical positional mappings remain one-based worksheet positions. Overlay validation rejects unknown or ignored properties, blank or duplicate effective headers, and conflicts with another active `MapFromHeader`. CellSharp accepts only final strings supplied by an application and implements no culture, translation, localization, or alias-dictionary logic.

## Multi-sheet workbooks

`ExcelWorkbookBuilder` is a small typed definition of ordered worksheets. `AddSheet(rows, schema, overlay, options)` and `AddTemplateSheet(schema, overlay, options)` accept an optional immutable overlay while preserving the simpler existing overloads. Duplicate names are rejected case-insensitively before a file is created. The single-sheet writers use the same worksheet writer and workbook-level style/validation components, avoiding separate conversion or presentation pipelines.

`ExcelWorkbookReader` owns an opened package and exposes `Read(schema)` by schema sheet name plus `ReadAt(index, schema)` for zero-based public-sheet position. The index method verifies the schema name to prevent accidental mismatches. Internal validation lookup sheets are hidden, excluded from indexes, and rejected as application data; explicitly named hidden user worksheets remain readable. One workbook-level data-validation pass creates at most one hidden lookup worksheet and deduplicates equal allowed-value lists.

The schema columns are transformed into the same internal import/export metadata used by conventions. There is no separate Open XML or reflection pipeline for schemas, so diagnostics and cell conversion semantics are unchanged.

## Generated XLSX templates

`Excel.CreateTemplate(path, schema)` creates a header-only workbook from an existing schema. It deliberately names an operation (`CreateTemplate`) rather than introducing an `ExcelTemplate` model, which would be confused with `ExcelStyleTemplate`. It reuses the export metadata, header generation, worksheet layout, style catalog, width calculation, and freeze behavior used by `Write`.

The template writes a default XLSX style on every column, rather than creating placeholder rows. This lets spreadsheet applications apply schema format and alignment to cells users add manually while keeping the document genuinely empty. Optional columns remain in the recommended complete shape; ignored columns remain excluded.

## Validation

Schema validation is column-level and runs only after a cell has converted successfully. There are two deliberately separate models. `Validate(...)` stores an arbitrary .NET delegate; it runs only at import time, receives the resulting value including `null`, and its exceptions propagate because they indicate application logic or configuration failures.

`AllowedValues(...)`, `Range(...)`, and `DateBetween(...)` store one structured internal declarative rule per column. It runs before custom delegates, returns `ValidationFailed` with a deterministic message, and still permits every applicable custom rule to run afterwards. Null values bypass numeric/date range checks; nullability remains a separate conversion concern.

The export and template writers share one native-Data-Validation component. It applies constraints to row 2 through row 1,048,576 without creating cells. Numeric and date ranges translate directly. Lists use a deterministic hidden `_CellSharpValidation` worksheet and a range formula instead of inline formulas, avoiding comma escaping and length limits. The main worksheet remains first, so import behavior is unchanged. Arbitrary delegates are intentionally never inferred as Excel rules.

## Custom converters

`ConvertWith<TCellValue>(IExcelValueConverter<TValue, TCellValue>)` attaches a reusable bidirectional converter to one schema column. It maps a logical property type to a declared CellSharp-supported scalar type and contains no reflection, workbook, cell, or Open XML knowledge. Both generic parameters are intentional: they make the logical property type and scalar contract explicit in the converter interface while ordinary inference keeps `ConvertWith` concise at the call site. This is the stabilized pre-1.0 converter API. The schema stores an internal adapter once, so no converter discovery or reflection is performed per cell.

On read, CellSharp first converts the raw cell to the declared scalar type, then invokes `TryRead`; a `false` result is an `InvalidValue` diagnostic and a converter exception propagates. On write it invokes `Write` first and applies the normal scalar writer, style, and format afterwards. Null cells bypass a converter. This preserves the pipeline `scalar conversion → converter → declarative validation → custom validation` and keeps non-native property types available only when explicitly configured with a converter.

Formula-looking strings are always emitted as inline text cells, not formula cells. This applies to headers, ordinary strings, string converter output, and list-validation values. It is an XLSX typed-cell guarantee rather than CSV-style prefix escaping; CellSharp does not deliberately alter the supplied text by prepending apostrophes. Formula exports store the property's converted value as their initial cache, and formula cells encountered during import use that cached value. Error cells and formulas without a cache are classified as `InvalidValue` data errors, not reader exceptions.

## Styling

Schema describes data structure; write options describe workbook presentation. A schema column may carry small export hints (`Format`, `Width`, and `Align`) because they are attached to a stable known field; they are ignored by the import pipeline and never change values. `ExcelTheme` is a shortcut that resolves internally into an `ExcelStyleTemplate`. A custom immutable `ExcelStyleTemplate` replaces the theme as the base visual identity and has no Open XML or workbook state, so one static instance can be safely reused across exports.

Precedence is column format/alignment, then explicit `HeaderStyle` values, then the custom template or resolved built-in theme, then fallback defaults. Explicit column width wins over write-level autofit estimation. Alternating rows obtain their color from the active base template instead of from writer logic.

The Open XML catalog deduplicates fonts, fills, borders, cell formats, and equivalent custom number formats. This keeps large exports compact while preserving a backend-independent public model.

Future cell styling should be a narrow declarative option with an export-cell context (row index, column metadata, value) and an immutable style delta. It can then compose after template and column styles, with a separate conditional-formatting API for rules that should remain native XLSX. No public placeholder has been introduced before that contract is needed.

## Deliberate limits

The vertical slice supports `string`, numeric primitives, `bool`, `DateTime`, `Guid`, nullable value-type variants, typed schema configuration, custom converters, custom and declarative column validation, visual templates, schema-generated XLSX templates, native formulas, caller-owned seekable streams, worksheet layout/styling, conditional formatting, defined names, comments, images, and report convenience components. It does not support constructor binding, header aliases, dynamic objects, complex properties without a converter, row validation, a formula engine, charts, themes, or structural row/column insertion.

## Evolution constraints

Adding overloads and option objects is additive; changing default header naming, schema worksheet naming, optional-column behavior, conversion semantics, schema inclusiveness, or the handling of invalid rows would be observable breaking changes. Future schema and validation features can refine value requirements and mapping while preserving the diagnostic model. Streaming and async APIs can be added beside the current synchronous API without changing it.

## API direction

`Read<T>` and typed schemas are available in this version. Future schema capabilities remain design direction only:

```csharp
Excel.Write("customers.xlsx", customers);

var result = Excel.Read<Customer>("customers.xlsx");

var schema = Excel.Schema<Customer>()
    .Column(x => x.Id)
    .Column(x => x.Name, column => column.Header("Customer Name"))
    .Build();
```
