# Writing XLSX files

Use `Excel.Write` to export a sequence of typed objects to the first worksheet of a new XLSX file:

```csharp
Excel.Write("customers.xlsx", customers);
```

The output path must name an existing directory. An existing file at that path is replaced. An empty sequence still creates a workbook with its header row.

## Conventional export

Without a schema, public readable non-indexer scalar properties become columns in their reflection order. Property names are headers.

Supported values are `string`; `byte`, `sbyte`, `short`, `ushort`, `int`, `uint`, `long`, `ulong`, `float`, `double`, and `decimal`; `bool`; `DateTime`; `Guid`; and nullable value-type equivalents. Null values become blank cells.

## Streams

All `Write` overloads also accept a `Stream`, including schemas, overlays, and presentation configuration:

```csharp
using var stream = new MemoryStream();

Excel.Write(stream, products, schema, overlay);
return stream.ToArray();
```

The stream must be writable and seekable. CellSharp resets its position to zero and truncates it before creating the workbook, then leaves it open with `Position == Length`. This prevents trailing bytes when a smaller workbook replaces a larger one. CellSharp writes directly to normal seekable streams and does not create an intermediate full-workbook buffer.

## Schemas

Use a schema to define the exported document shape explicitly:

```csharp
public sealed class Customer
{
    public int Id { get; set; }
    public string? Name { get; set; }
    public string? InternalCode { get; set; }
}

var schema = Excel.Schema<Customer>()
    .SheetName("Customers")
    .Column(x => x.Name, column => column.Header("Customer Name"))
    .Column(x => x.Id, column => column.Header("Customer ID"))
    .Column(x => x.InternalCode, column => column.Ignore())
    .Build();

Excel.Write("customers.xlsx", customers, schema);
```

Only declared, non-ignored properties are exported. Declaration order is column order. The schema property must be public, readable, writable, and of a supported scalar type. See [schemas](schemas.md) for the shared read/write behavior.

`SheetName(...)` on a schema controls the worksheet name for this export and for matching reads and generated templates. Without it, generated workbooks retain the `Sheet1` default.

## Attribute-based schemas

For a stable DTO mapping, declare the basic column metadata on the model and opt in explicitly:

```csharp
public sealed class Customer
{
    [ExcelColumn("Customer ID", Order = 1)]
    public int Id { get; set; }

    [ExcelColumn("Customer name", Order = 2, Optional = true)]
    public string? Name { get; set; }

    [ExcelIgnore]
    public string? InternalCode { get; set; }
}

var schema = Excel.SchemaFromAttributes<Customer>();
Excel.Write("customers.xlsx", customers, schema);
```

`ExcelColumn` supports a header, `Optional`, an Excel `Format`, and `Order`; `ExcelIgnore` excludes a property. Properties without either attribute still participate with their property name as the header. Explicitly ordered properties come first in ascending order; ties and unordered properties retain reflection discovery order. `Order` is optional—leaving it unset is distinct from assigning any integer value.

Attributes have no effect on `Excel.Write(path, rows)`, `Excel.Read<T>(path)`, or any other schema-less API. They are read only by `Excel.SchemaFromAttributes<T>()`.

For a call-site override or features that do not fit attributes, configure the generated schema fluently. A configured property replaces its attribute defaults rather than creating a second column:

```csharp
var schema = Excel.SchemaFromAttributes<Customer>(builder => builder
    .Column(x => x.Name, column => column
        .Header(applicationLabel)
        .MapFromHeader("Imported name")));
```

Converters, formulas, validation, styling, and other dynamic behavior remain fluent-only.

For multiple typed worksheets in one file, use `Excel.Workbook().AddSheet(...).Write(path)`. Each sheet supplies its own schema and optional presentation configuration; the existing `Excel.Write` overloads remain the single-sheet API. See [multi-sheet workbooks](multi-sheet.md).

## Native Excel Tables

Opt in to a native Excel Table, including its header filters, with `AsTable(...)` on the schema:

```csharp
var schema = Excel.Schema<Customer>()
    .SheetName("Customers")
    .AsTable("CustomersTable")
    .Column(x => x.Id)
    .Column(x => x.Name)
    .Build();
```

Tables are export metadata only: the normal worksheet values, formats, validation, and formula cells are still written by CellSharp's existing pipeline. See [Excel Tables](tables.md) for naming, styles, templates, and multi-sheet behavior.

## Per-operation headers and columns

Use an immutable overlay when the effective labels or enabled modules are known only at operation time:

```csharp
var overlay = Excel.Overlay<Customer>(runtime => runtime
    .Header(x => x.Name, applicationLabel)
    .Include(x => x.InternalCode, moduleEnabled));

Excel.Write("customers.xlsx", customers, schema, overlay);
```

Disabled columns are omitted without blank placeholders; active columns keep schema order. `applicationLabel` is a string produced by the application—CellSharp does not localize it. Formatting, widths, converters, and declarative validation apply only to active columns.

## Themes, templates, and header styling

Every export uses the restrained `ExcelTheme.Default` presentation: a readable header, base cell formatting, and compact shared style records. Choose another built-in theme or override selected header values with write options:

```csharp
Excel.Write("customers.xlsx", customers, options => options
    .Theme(ExcelTheme.Modern)
    .HeaderStyle(style => style
        .Bold()
        .Background("#202020")
        .Foreground("#FFFFFF"))
    .AutoFitColumns()
    .AlternatingRows()
    .FreezeHeaderRow());
```

`AutoFitColumns()` estimates widths from exported values; it is not Excel's exact application-side autofit. `AlternatingRows()` uses the selected theme or template's alternate color, and `FreezeHeaderRow()` freezes row one. A schema column's explicit `Width(...)` takes precedence over autofit. See [styling](styling.md) for templates, format codes, and limits.

For schema-owned filters, configurable frozen panes, and essential print setup, see [worksheet settings](worksheet-settings.md). `FreezeHeaderRow()` remains the operation-level compatibility shortcut for freezing row one.

For a header-only XLSX file generated from the same schema rather than from rows, see [generated templates](templates.md).

When a supplied schema uses declarative validation, `Write` also emits the equivalent native Excel Data Validation for future edits of the workbook. Custom `.Validate(...)` predicates remain runtime-only. See [validation](validation.md).

A schema column can use `ConvertWith` to turn a domain value into a supported cell scalar before writing. `Format(...)` then applies to that scalar representation. See [custom converters](converters.md).

## Text and formulas

Strings are always written as text. For example, this is not written as an Excel formula:

```csharp
new Customer { Name = "=SUM(A1:A2)" }
```

To intentionally write native Excel formula cells, configure a schema column with `Formula(...)`. See [formulas](formulas.md). CellSharp does not create charts, macros, `DataTable`, or JSON workbooks. For imports, see [reading](reading.md); for supported target frameworks, see [compatibility](compatibility.md).

This is XLSX typed-cell behavior, not CSV escaping: strings are serialized as inline string cells and never as formula cells, including values beginning with `=`, `+`, `-`, or `@`. The same applies to string converter output, generated headers, and native-validation list values. CellSharp does not prepend a destructive apostrophe.
