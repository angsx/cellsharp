# Typed schemas

A schema describes a stable XLSX representation for one .NET type and is used for both writing and reading.

```csharp
public sealed class Customer
{
    public int Id { get; set; }
    public string? Name { get; set; }
    public string? InternalCode { get; set; }
}

var schema = Excel.Schema<Customer>()
    .SheetName("Customers")
    .Column(x => x.Id, column => column.Header("Customer ID"))
    .Column(x => x.Name, column => column.Header("Customer Name").Optional())
    .Column(x => x.InternalCode, column => column.Ignore())
    .Build();

Excel.Write("customers.xlsx", customers, schema);
var result = Excel.Read<Customer>("customers.xlsx", schema);
```

## Worksheet name

`SheetName(...)` is the single schema-level setting for the worksheet used by `Write`, `Read`, and `CreateTemplate`. Its default is `Sheet1`, preserving the conventional behavior. A schema with a configured name reads that worksheet by name, case-insensitively, and throws `InvalidOperationException` if it is absent; it never falls back to another sheet.

The same schema name identifies its worksheet in `Excel.Workbook()` export/template definitions and in `Excel.Open(...).Read(schema)`. Workbook sheet names must be unique case-insensitively. See [multi-sheet workbooks](multi-sheet.md).

Excel worksheet names must be non-empty, no more than 31 characters, and cannot contain control characters or `:`, `\`, `/`, `?`, `*`, `[`, or `]`. CellSharp writes valid names exactly as supplied; it does not escape or rewrite them.

## Inclusive document shape

Schemas are inclusive: only properties declared with `Column` participate. Properties not declared are left out, so adding a property to `Customer` does not silently change an established XLSX contract. Declared order is the export column order.

`Header` changes the exported header and the header expected during import. Matching remains case-insensitive. `Ignore` excludes the declared property from both directions.

## Runtime import mapping

Use a column configuration to decouple an incoming workbook layout from the schema's generated header:

```csharp
var schema = Excel.Schema<Product>()
    .Column(x => x.Color, column => column
        .Header("Product color")
        .MapFromColumn(3))
    .Column(x => x.Size, column => column.MapFromHeader("size"))
    .Build();
```

`MapFromColumn(int)` uses one-based Excel column numbers: `1` is A, `2` is B, and so on. `MapFromHeader(string)` matches case-insensitively. They affect only import; the schema's `Header(...)` remains the write/template header.

Precedence is deterministic: `MapFromColumn`, then an active runtime overlay header, then `MapFromHeader`, then normal `Header` matching. The winning mapping continues through the ordinary conversion, converter, declarative validation, and `.Validate(...)` pipeline. Required mappings that cannot be found report `MissingHeader`; `Optional()` mappings may be absent. Ignored columns cannot have a runtime mapping.

The builder rejects duplicate runtime positions, duplicate runtime headers, and a runtime header that would compete with another property's schema header.

## Runtime schema overlay

`ExcelSchemaOverlay<T>` is an immutable per-operation layer over an `ExcelSchema<T>`. Build one with `Excel.Overlay<T>(...)` and pass it to `Write`, `Read`, `CreateTemplate`, or the multi-sheet APIs:

```csharp
var overlay = Excel.Overlay<Product>(runtime => runtime
    .Header(x => x.Size, translatedSizeHeader)
    .Include(x => x.SerialNumber, moduleEnabled));

Excel.Write("products.xlsx", products, schema, overlay);
```

`Header` supplies the effective column header for that operation: it is written during export/template generation and expected during import. It does not mutate `schema`. `MapFromHeader` remains a schema-owned import fallback for a different incoming layout; it is never written. `MapFromColumn` remains a one-based physical Excel position and is not renumbered when other columns are disabled.

`Include(..., false)` removes a column from the operation. Export and templates compact the remaining active columns while preserving their relative order. Import ignores both presence and absence of the disabled column: it creates no `MissingHeader`, converter, validation, or native Excel Data Validation work. `Optional()` instead keeps a column active and only permits its incoming header to be absent. `Ignore()` excludes a property permanently, and configuring it in an overlay fails fast.

An active optional column is still converted and validated when it is present. Consequently an invalid value in an `Optional()` column produces the same structured import diagnostic as a required column, while a runtime-disabled column produces none.

Overlay setup validates blank headers, unknown or ignored properties, duplicate effective headers (case-insensitively), and collisions between an effective runtime header and another column's `MapFromHeader`. The same schema and separate overlays can be used concurrently because neither object contains mutable operation state. CellSharp accepts final strings supplied by application code; it has no culture, translation, resource-manager, fallback, or alias-dictionary API.

## Validation

Use `Validate` to attach a predicate to the converted value of one column:

```csharp
.Column(x => x.Age, column => column
    .Validate(age => age >= 18, "Customer must be at least 18"))
```

Rules run after conversion and before a row is added to `Rows`. A failed rule produces `ValidationFailed`; a conversion failure produces `InvalidValue` and does not run the rule. Multiple rules are evaluated in declaration order, and every failed rule is reported. Predicate exceptions propagate to the caller rather than being hidden as invalid workbook data.

Structured rules are also available through `AllowedValues(...)`, `Range(...)`, and `DateBetween(...)`. Unlike arbitrary `Validate(...)` delegates, they can become native Excel Data Validation in generated files. See [validation](validation.md) for the supported types and the separation between the two models.

## Custom conversion

`ConvertWith<TCellValue>(IExcelValueConverter<TValue, TCellValue>)` maps a property to and from a supported scalar representation. It permits a non-native domain type in an otherwise typed schema and runs before validation during import. The logical property type and cell scalar type are intentionally both explicit, keeping converters type-safe and discoverable. See [custom converters](converters.md).

## Required and optional columns

Declared columns are required by default. `Optional()` allows the entire column header to be absent during import; its property then keeps the normal .NET default value.

This rule concerns the presence of a column, not the value in each cell. Cell nullability remains separate: a nullable value-type property can receive a blank cell, while a blank non-nullable value-type cell is an import error.

## Presentation hints for export

Schemas remain primarily structural, but a column can also carry small export-only presentation hints. They never change the .NET value, import conversion, or validation behavior:

```csharp
var schema = Excel.Schema<Order>()
    .Column(x => x.Date, column => column
        .Format("dd/MM/yyyy")
        .Width(14))
    .Column(x => x.Total, column => column
        .Format("#,##0.00")
        .Align(ExcelHorizontalAlignment.Right))
    .Build();
```

`Format` accepts an Excel format-code string and writes it to the XLSX style table; CellSharp does not parse it. `Width` must be greater than zero and no greater than 255. `Align` accepts the deliberately small `General`, `Left`, `Center`, and `Right` enum. These settings are consumed by `Write` and `CreateTemplate`; the same schema stays reusable for `Read`. See [styling](styling.md) and [generated templates](templates.md) for precedence and workbook-level options.

`Formula(...)` is a separate write-time setting: it produces a native formula cell for each data row. The property's converted value is stored as that cell's initial cached result, while the formula remains authoritative when a spreadsheet recalculates it. It can be combined with format, width, alignment, converters, and custom `.Validate(...)` predicates, but not native declarative validation. See [formulas](formulas.md).

## Build and reuse

`Build()` creates an immutable `ExcelSchema<T>`. A schema can be reused for multiple reads, writes, and generated templates and contains no state from an individual operation. `ExcelSchemaOverlay<T>` is likewise immutable and can be reused independently.

Schema configuration errors are reported early. Each configured member must be a direct public readable and writable scalar property. Duplicate properties, duplicate effective headers, unsupported types, invalid expressions, and empty schemas are rejected during construction.

For the conventional behavior without a schema, see [writing](writing.md) and [reading](reading.md). For first use, see [getting started](getting-started.md).
