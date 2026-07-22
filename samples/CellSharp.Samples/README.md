# CellSharp Samples

Run all scenarios from a repository checkout:

```bash
dotnet run --project samples/CellSharp.Samples
```

The runner recreates `samples/CellSharp.Samples/output/` and prints every generated workbook. Open those XLSX files in Excel or LibreOffice, change a scenario, and run it again.

| Scenario | Workbook | Learn more |
|---|---|---|
| Basic export | `basic-export.xlsx` | [Export a list](../../docs/use-cases/export-a-list.md) |
| Basic import | `basic-import.xlsx` | [Import a list](../../docs/use-cases/import-a-list.md) |
| Fluent schema | `fluent-schema.xlsx` | [Reusable domain schema](../../docs/use-cases/reusable-domain-schema.md) |
| Attribute schema | `attribute-schema.xlsx` | [Attribute-based schema](../../docs/use-cases/attribute-schemas.md) |
| Custom converter | `custom-converter.xlsx` | [Custom converters](../../docs/use-cases/custom-converters.md) |
| Data-entry template | `product-import-template.xlsx` | [Data-entry workbook](../../docs/use-cases/data-entry-workbook.md) |
| Import round trip | `product-import-round-trip.xlsx` | [Import and validate](../../docs/use-cases/import-and-validate.md) |
| Multi-sheet workbook | `customer-orders.xlsx` | [Multi-sheet workbook](../../docs/use-cases/multi-sheet-workbook.md) |
| Formula columns | `formula-columns.xlsx` | [Formulas and calculated columns](../../docs/use-cases/formulas-and-calculated-columns.md) |
| Report and table | `sales-report.xlsx` | [Generate a report](../../docs/use-cases/generate-a-report.md) |
| Streams | `stream-export.xlsx` | [Streams and web APIs](../../docs/use-cases/streams-and-web-apis.md) |

The code is deliberately small and lives in [Scenarios](Scenarios). Feature details stay in the linked documentation.
