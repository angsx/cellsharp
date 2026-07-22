# Changelog

## 0.6.0-prerelease

- Prepared the public prerelease with strongly typed read/write APIs, schemas, validation, formulas, templates, tables, images, report components, and multi-sheet workbooks.
- Added resilience, security, diagnostics, styling, stream, and documentation improvements for real-world XLSX use.

## Unreleased

### Added

- Added per-read fallback culture support for numeric and date values stored as text, plus opt-in `EmptyStringAsNull` normalization for explicit empty text cells.
- Added bounded XLSX package, Open XML part, worksheet row, and diagnostic limits for untrusted input, plus configurable `ExcelReadOptions` limits and `Excel.Open(..., options)` overloads.
- Added security regression tests, O(1) shared-string and date-style lookup, physical-column validation, and explicit XLSX row/column output bounds.
- Path reads now size-check and parse the same open file handle, preventing a path replacement from bypassing package limits.
- Added NuGet audit enforcement, reproducible package lock files, Dependabot configuration, least-privilege CI permissions, and immutable GitHub Action pins.
- Updated the pinned .NET 8 SDK to the current security-serviced 8.0.423 release.
- Added post-conversion schema validation and `ValidationFailed` read errors.
- Added built-in export themes, header style overrides, estimated column widths, and frozen header rows.
- Added immutable reusable `ExcelStyleTemplate` instances, alternating data rows, and template-based theme resolution.
- Added schema-column Excel format codes, explicit widths, and horizontal alignment with style and number-format deduplication.
- Added `Excel.CreateTemplate` for generating empty, styled XLSX templates from reusable schemas.
- Corrected serialized style colors and font child ordering for schema-valid OpenXML workbooks.
- Added declarative `AllowedValues`, `Range`, and `DateBetween` schema validation for import and native Excel Data Validation in writes and generated templates.
- Added reusable bidirectional `IExcelValueConverter<TValue, TCellValue>` support for schema columns.
- Added schema-level `SheetName(...)`, shared by `Write`, `Read`, and `CreateTemplate`, with explicit missing-sheet failures on read.
- Stabilized `IExcelValueConverter<TValue, TCellValue>` and `ConvertWith<TCellValue>(...)` as the pre-1.0 converter API.
- Added resilient import handling for styled-only rows, sparse cells, unreliable worksheet dimensions, Excel error cells, and formulas with or without cached values.
- Added schema-column runtime import mapping through `MapFromColumn(...)` and `MapFromHeader(...)`.
- Documented and regression-tested literal XLSX text handling for formula-looking strings.
- Added typed multi-sheet export, import, and template generation through `Excel.Workbook()` and `Excel.Open(...)`.
- Added immutable per-operation `ExcelSchemaOverlay<T>` configuration for runtime header values and conditional column participation across single-sheet, template, and multi-sheet APIs.
- Added structured import error categories and CLR property metadata, plus immutable per-read invalid-row materialization policy through `ExcelReadOptions`.
- Added first-class seekable stream overloads for read, write, templates, multi-sheet workbooks, and `Excel.Open`, preserving caller stream ownership.
- Added cooperative `CancellationToken` support for selected synchronous stream reads, writes, templates, and multi-sheet writes; no artificial async APIs were introduced.
- Added native schema-column formula export with one-based Excel row context, formula normalization, and workbook recalculation hints; cached formula values continue through the normal import pipeline.
- Added opt-in native Excel Table export through `AsTable(...)`, including deterministic global names, native filters, built-in styles, templates, and multi-sheet support.
- Added schema-level standalone AutoFilter, configurable frozen panes, and essential native print settings including orientation, fit-to-page, gridlines, and repeated header rows.

### Breaking changes

- Removed `Required()`. Declared schema columns are required by default; use `Optional()` when an import header may be absent.
- Changed the public column builder to `ExcelColumnBuilder<T, TValue>` so validation predicates are strongly typed.

## 0.1.0-alpha.1

- Added typed object export to a single XLSX worksheet.
