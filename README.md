# CellSharp

> Strongly typed Excel for .NET without exposing the spreadsheet plumbing.

CellSharp is a lightweight, strongly typed .NET library for reading, writing, and validating Excel XLSX files. Use conventions for simple exports, then introduce schemas when the workbook becomes a contract.

Define your schema once. CellSharp handles mapping, conversion, validation, workbook generation and the boring spreadsheet stuff for you.

```csharp
using CellSharp;

public sealed class Customer
{
    public int Id { get; set; }
    public string Name { get; set; } = "";
    public DateTime CreatedAt { get; set; }
}

var customers = new[]
{
    new Customer { Id = 1, Name = "Acme Ltd", CreatedAt = new DateTime(2025, 1, 15) },
    new Customer { Id = 2, Name = "Contoso", CreatedAt = new DateTime(2025, 2, 8) },
};

Excel.Write("customers.xlsx", customers);
```

Start with conventions, then introduce a schema when the workbook becomes a contract.

## Install

CellSharp `0.6.0-prerelease` is prerelease software; API details may still evolve before `1.0`:

```bash
dotnet add package CellSharp --version 0.6.0-prerelease
```

The package ID is `CellSharp`.

## Read with a schema

Use a schema when column names and worksheet names are part of the file contract:

```csharp
var schema = Excel.Schema<Customer>()
    .SheetName("Customers")
    .Column(x => x.Id, column => column.Header("Customer ID"))
    .Column(x => x.Name, column => column.Header("Customer Name"))
    .Build();

var result = Excel.Read<Customer>("customers.xlsx", schema);
```

See [Getting started](https://github.com/angsx/cellsharp/blob/main/docs/getting-started.md) for the complete write, read, and schema workflow.

## What it does

- Typed import and export, with reusable inclusive schemas and configurable fallback culture for text-form numbers and dates.
- Structured import diagnostics and `Skip`/`Include` invalid-row policies.
- Native Excel validation, templates, formulas, Tables, conditional formatting, named ranges, comments, headers/footers, page breaks, PNG/JPEG images, filters, panes, report components, and essential print settings.
- Multi-sheet workbooks, runtime schema overlays, custom scalar converters, and caller-owned seekable streams.
- Cooperative cancellation for selected synchronous stream operations; no artificial async API.

Normal strings are always exported as text, even when they begin with `=`. Formulas require an explicit `Formula(...)` column configuration.

## What can I build with CellSharp?

- Export typed objects with zero configuration
- Build reusable fluent or attribute-based schemas
- Generate Excel data-entry templates with native validation
- Import and diagnose user-edited workbooks
- Create multi-sheet workbooks, formulas, and native Excel Tables
- Stream XLSX files directly from a web application

See the task-oriented [use cases](https://github.com/angsx/cellsharp/blob/main/docs/use-cases/README.md), or clone the repository and run the [CellSharp Samples](https://github.com/angsx/cellsharp/blob/main/samples/CellSharp.Samples/README.md).

## Simple when you need it

```csharp
Excel.Write("customers.xlsx", customers);
```

## Powerful when you need more

```csharp
Excel.Workbook().AddSheet(sales, schema, sheet =>
{
    sheet.Title("Sales Performance", "B1:F2");
    sheet.Kpi("Revenue", 125_400m, "A5:B7").Value.NumberFormat("€ #,##0");
    sheet.Range("B11:B200").ConditionalFormat().GreaterThan(10000).Style(s => s.FillColor("#C6EFCE"));
}).Write("sales-report.xlsx");
```

## Quick import

```csharp
var result = Excel.Read<Customer>("customers.xlsx");

foreach (var error in result.Errors)
{
    Console.WriteLine($"{error.SheetName} {error.CellReference}: {error.Message}");
}
```

`Rows` contains valid rows by default. Pass `new ExcelReadOptions(ExcelInvalidRowPolicy.Include)` to retain partial invalid rows for remediation.

Imports use invariant XLSX conventions by default. For third-party files that store localized numbers or dates as text, pass a fallback `CultureInfo`; `EmptyStringAsNull` can also map explicit empty text cells to `null` for strings and nullable value types. See [reading XLSX files](https://github.com/angsx/cellsharp/blob/main/docs/reading.md#localized-text-values) for the exact precedence and nullability rules.

## Documentation

- [Getting started](https://github.com/angsx/cellsharp/blob/main/docs/getting-started.md)
- [Use cases](https://github.com/angsx/cellsharp/blob/main/docs/use-cases/README.md): complete tasks from export to web streams
- [Writing XLSX files](https://github.com/angsx/cellsharp/blob/main/docs/writing.md) and [reading XLSX files](https://github.com/angsx/cellsharp/blob/main/docs/reading.md)
- [Schemas and overlays](https://github.com/angsx/cellsharp/blob/main/docs/schemas.md), [validation](https://github.com/angsx/cellsharp/blob/main/docs/validation.md), [diagnostics](https://github.com/angsx/cellsharp/blob/main/docs/diagnostics.md), and [custom converters](https://github.com/angsx/cellsharp/blob/main/docs/converters.md)
- [Templates](https://github.com/angsx/cellsharp/blob/main/docs/templates.md), [formulas](https://github.com/angsx/cellsharp/blob/main/docs/formulas.md), [Excel Tables](https://github.com/angsx/cellsharp/blob/main/docs/tables.md), and [worksheet settings](https://github.com/angsx/cellsharp/blob/main/docs/worksheet-settings.md)
- [Conditional formatting](https://github.com/angsx/cellsharp/blob/main/docs/conditional-formatting.md)
- [Named ranges](https://github.com/angsx/cellsharp/blob/main/docs/named-ranges.md) and [worksheet utilities](https://github.com/angsx/cellsharp/blob/main/docs/worksheet-settings.md)
- [Images](https://github.com/angsx/cellsharp/blob/main/docs/images.md)
- [Report layouts](https://github.com/angsx/cellsharp/blob/main/docs/report-layouts.md)
- [AI agent guide](https://github.com/angsx/cellsharp/blob/main/docs/ai-agent-guide.md): concise contracts and canonical workflows for coding agents and contributors
- [Streams and cancellation](https://github.com/angsx/cellsharp/blob/main/docs/streams.md) and [multi-sheet workbooks](https://github.com/angsx/cellsharp/blob/main/docs/multi-sheet.md)
- [Security model and untrusted XLSX input](https://github.com/angsx/cellsharp/blob/main/docs/security.md)
- [Release checklist](https://github.com/angsx/cellsharp/blob/main/docs/release-security.md)
- [Roadmap](https://github.com/angsx/cellsharp/blob/main/ROADMAP.md)
- Recipes: [validated import](https://github.com/angsx/cellsharp/blob/main/docs/recipes/import-with-validation.md), [order template](https://github.com/angsx/cellsharp/blob/main/docs/recipes/create-an-order-template.md), [multi-sheet workbook](https://github.com/angsx/cellsharp/blob/main/docs/recipes/multi-sheet-workbook.md), [formulas and Tables](https://github.com/angsx/cellsharp/blob/main/docs/recipes/formulas-and-tables.md)

## Compatibility and scope

CellSharp targets `netstandard2.0` and `net8.0`. It intentionally does not provide charts, table resizing, totals rows, advanced print layout, or a formula engine.

## Contributing

Contributions are welcome. See [CONTRIBUTING.md](https://github.com/angsx/cellsharp/blob/main/CONTRIBUTING.md) for development, testing, and pull request guidance.

Please report security issues according to [SECURITY.md](https://github.com/angsx/cellsharp/blob/main/SECURITY.md).

## Support CellSharp

CellSharp is developed and maintained in my spare time. If it saves you time or you simply enjoy using it, you can support its continued development on Ko-fi.

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/W8O623M82T)

Support is entirely optional and does not affect access to the project, its features, or technical decisions.

## License

CellSharp is licensed under the MIT License. See [LICENSE](https://github.com/angsx/cellsharp/blob/main/LICENSE) for details.
