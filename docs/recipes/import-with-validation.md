# Recipe: import with validation

Use declarative rules when a constraint belongs both in import and in a generated template. Use `Include` when an operator must review malformed rows.

```csharp
var schema = Excel.Schema<Product>()
    .SheetName("Products")
    .Column(x => x.Name)
    .Column(x => x.Category, column => column.AllowedValues("Retail", "Wholesale"))
    .Column(x => x.Quantity, column => column.Range(1, 1000))
    .Build();

var result = Excel.Read(
    "uploaded-products.xlsx",
    schema,
    new ExcelReadOptions(ExcelInvalidRowPolicy.Include));

foreach (var error in result.Errors)
{
    Console.WriteLine($"{error.CellReference}: {error.Message}");
}
```

The result retains reviewable partial rows while keeping `IsValid` false until every data issue is fixed.

For a template-to-import workflow and executable round trip, see [Import and validate](../use-cases/import-and-validate.md).
