# Getting started

Install the released package from a configured NuGet source:

```bash
dotnet add package CellSharp
```

For local package verification, pack CellSharp and add that folder as a NuGet source; consume `CellSharp` with `PackageReference`, not a project reference. The package remains pre-1.0 until a stable release is announced.

## First export

```csharp
using CellSharp;

public sealed class Customer
{
    public int Id { get; set; }
    public string? Name { get; set; }
}

var customers = new[]
{
    new Customer { Id = 1, Name = "Alice" },
    new Customer { Id = 2, Name = "Bob" },
};

Excel.Write("customers.xlsx", customers);
```

CellSharp creates a workbook with an `Id` and `Name` header row followed by the customer values.

This is the first step in the [progressive use-case path](use-cases/README.md): start with conventions, introduce a schema when the workbook shape becomes a contract, then add validation and templates when users edit files.

## First import

```csharp
var result = Excel.Read<Customer>("customers.xlsx");

if (result.IsValid)
{
    foreach (var customer in result.Rows)
    {
        Console.WriteLine(customer.Name);
    }
}
else
{
    foreach (var error in result.Errors)
    {
        Console.WriteLine($"{error.CellReference}: {error.Message}");
    }
}
```

`Rows` contains only fully converted rows. Invalid rows are not materialized and are described in `Errors`.

For a complete upload workflow, see [Import and validate](use-cases/import-and-validate.md).

## Diagnose imperfect files

Every import issue contains one-based Excel coordinates when a cell is involved, the CLR property, a programmatic category, the raw value, and a readable message:

```csharp
foreach (var error in result.Errors)
{
    Console.WriteLine($"Row {error.Row}, {error.PropertyName}: {error.Message}");
}
```

The default read policy is `ExcelInvalidRowPolicy.Skip`, preserving the normal behavior above. A remediation import can keep invalid rows with their successfully converted values:

```csharp
var review = Excel.Read<Customer>(
    "customers.xlsx",
    schema,
    new ExcelReadOptions(ExcelInvalidRowPolicy.Include));
```

Failed conversions retain the CLR default (`null` or `default(TProperty)`) in an included row. `review.IsValid` is still `false` while any errors exist.

## First schema

Use a schema when the workbook shape is part of your contract:

```csharp
var schema = Excel.Schema<Customer>()
    .SheetName("Customers")
    .Column(x => x.Id, column => column.Header("Customer ID"))
    .Column(x => x.Name, column => column.Header("Customer Name").Optional())
    .Build();
```

Pass the same schema to `Write`, `Read`, or `CreateTemplate`. `SheetName("Customers")` makes that schema use the named worksheet in all three operations; omitting it keeps the compatible `Sheet1` default. See [schemas](schemas.md), then [writing](writing.md) and [reading](reading.md) for behavior details.

For multiple typed worksheets in one file, use `Excel.Workbook()` and then `Excel.Open(...)`; the single-sheet calls remain the normal first choice. See [multi-sheet workbooks](multi-sheet.md).

## Runtime business configuration

When a label or module is decided by the application at runtime, keep the schema reusable and pass an overlay:

```csharp
var overlay = Excel.Overlay<Customer>(runtime => runtime
    .Header(x => x.Name, applicationSuppliedName)
    .Include(x => x.InternalCode, moduleEnabled));

Excel.Write("customers.xlsx", customers, schema, overlay);
```

CellSharp receives final strings only; resolving cultures, resources, or translations remains application code.

## In-memory files

Every path-based read, write, template, and multi-sheet operation also accepts a seekable stream:

```csharp
using var stream = new MemoryStream();

Excel.Write(stream, customers, schema);
stream.Position = 0;

var imported = Excel.Read(stream, schema);
```

CellSharp leaves caller-owned streams open. Writes overwrite from the start and leave the position at the end; reads begin from the start. This suits HTTP upload/download code without adding web-framework dependencies.

For long synchronous operations, pass a final `CancellationToken` to the relevant stream overload. Cancellation is cooperative and throws `OperationCanceledException`; it does not produce import diagnostics. See [async and cancellation](async-and-cancellation.md).

## Generate a blank template

The same schema can create a header-only XLSX file ready for manual completion:

```csharp
Excel.CreateTemplate("customers-template.xlsx", schema);
```

The generated file includes all non-ignored schema columns, including columns marked `Optional()`. See [generated templates](templates.md) for style and validation limits.

## Declarative validation

Use structured rules when a constraint must work during import and in a generated template:

```csharp
var schema = Excel.Schema<Customer>()
    .Column(x => x.Status, column => column.AllowedValues("Active", "Inactive"))
    .Column(x => x.Age, column => column.Range(18, 120))
    .Build();
```

See [validation](validation.md) for the distinction from custom `.Validate(...)` predicates.

## Domain values

Use `ConvertWith` when a schema property is not one of CellSharp's native scalar types:

```csharp
.Column(x => x.Status, column => column.ConvertWith(StatusConverter))
```

The converter is used by both `Write` and `Read`. See [custom converters](converters.md).

## First themed export

```csharp
Excel.Write(
    "customers.xlsx",
    customers,
    options => options.Theme(ExcelTheme.Modern));
```

See [styling](styling.md) for built-in themes, reusable visual templates, column formatting, and header customization.
