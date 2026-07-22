# Use cases

Feature guides explain the API. These pages answer a different question: “I need to do this with Excel in .NET—where do I start?” Every major page links to an executable scenario in [CellSharp Samples](../../samples/CellSharp.Samples/README.md).

| I want to… | Example |
|---|---|
| Export a list of objects to Excel | [Export a list](export-a-list.md) |
| Import a straightforward workbook by property name | [Import a list](import-a-list.md) |
| Control headers, order, and formatting | [Reusable domain schema](reusable-domain-schema.md) |
| Put stable mapping metadata on a DTO | [Attribute-based schema](attribute-schemas.md) |
| Map domain values such as status codes | [Custom converters](custom-converters.md) |
| Let users fill in an Excel file | [Data-entry workbook](data-entry-workbook.md) |
| Import a user-edited workbook and show errors | [Import and validate](import-and-validate.md) |
| Build a workbook with multiple typed sheets | [Multi-sheet workbook](multi-sheet-workbook.md) |
| Create calculated columns | [Formulas and calculated columns](formulas-and-calculated-columns.md) |
| Produce a polished table-based report | [Generate a report](generate-a-report.md) |
| Read and write without temporary files | [Streams and web APIs](streams-and-web-apis.md) |

## A progressive path

1. Start with `Excel.Write(...)` or `Excel.Read<T>(...)`.
2. Add a fluent schema when labels, order, or formatting matter.
3. Use attributes for stable, reusable DTO mappings.
4. Add validation and generate templates.
5. Import with diagnostics, then move on to workbooks, formulas, Tables, and streams.
