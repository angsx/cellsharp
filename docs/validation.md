# Validation

CellSharp has two intentionally separate schema-validation models. Both run after a value has converted successfully and both report `ValidationFailed`, but only one can be represented in Excel.

## Runtime custom validation

`Validate(...)` accepts an arbitrary .NET predicate:

```csharp
.Column(x => x.Email, column => column
    .Validate(IsValidEmail, "Email is invalid"))
```

It runs during `Excel.Read<T>` only. All custom rules are evaluated, and a predicate exception propagates to the caller instead of being presented as bad workbook data. Because the predicate is arbitrary application code, CellSharp does not translate it to Excel.

## Declarative validation

Use structured rules for constraints that CellSharp can evaluate during import and write as native Excel Data Validation:

```csharp
var schema = Excel.Schema<Customer>()
    .Column(x => x.Status, column => column
        .AllowedValues("Active", "Inactive", "Pending"))
    .Column(x => x.Age, column => column.Range(18, 120))
    .Column(x => x.BirthDate, column => column.DateBetween(
        new DateTime(1900, 1, 1),
        new DateTime(2030, 12, 31)))
    .Build();
```

`AllowedValues` is available for string columns. Values are deduplicated in first-occurrence order; empty or null values and an empty list are rejected. `Range` supports CellSharp numeric column types, including nullable numeric types; it is inclusive and ignores a converted null. `DateBetween` supports `DateTime` and `DateTime?`, is inclusive, and also ignores null. Bounds are checked during configuration.

One declarative constraint can be configured per column. It runs before custom `Validate(...)` rules; then all applicable custom rules still run, so multiple errors may be reported for the same cell.

When `ConvertWith(...)` is configured, both declarative and custom runtime validation receive the final logical property value. Native Excel validation is emitted only for combinations already meaningful for the schema's declared rule and cell representation; CellSharp does not infer a mapping between domain values and dropdown entries.

## Excel files and templates

The same structured rules are emitted for schema-based `Excel.Write(...)` and `Excel.CreateTemplate(...)`. They apply to rows 2 through 1,048,576 without adding blank cells. Allowed-value lists use an internal hidden worksheet so values containing commas, quotes, or long lists do not depend on fragile inline formulas.

```csharp
Excel.CreateTemplate("customers-template.xlsx", schema);
var result = Excel.Read<Customer>("customers-template.xlsx", schema);
```

Custom predicates are runtime-only; list, numeric range, and date range rules are the ones represented in the XLSX file. `Optional()` controls only whether a header may be absent during import. If that optional column is present, its declarative validation runs normally; it is also included and validated in generated templates. `Ignore()` cannot be combined with a declarative validation.

## Limits

Native Excel validation does not make a field required, and it does not replace CellSharp's conversion or nullability rules. Dropdowns, numeric ranges, and dates are supported; regexes, cross-property constraints, custom delegates, and conditional formatting are not.

For the complete generate-edit-import workflow, see [Data-entry workbook](use-cases/data-entry-workbook.md) and [Import and validate](use-cases/import-and-validate.md).
