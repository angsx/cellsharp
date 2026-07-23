# Formulas

CellSharp writes formulas as native XLSX formula cells. It does not parse, calculate, or validate formulas. The current property value is written as the formula's cached result, so a newly exported workbook can be imported immediately; Excel, LibreOffice, or another calculation-capable consumer replaces that cache when it recalculates the formula.

## Writing formulas

Configure a formula on a schema column. The callback receives the real one-based Excel row, including the header row, so the first data row is `2`:

```csharp
var schema = Excel.Schema<Invoice>()
    .Column(x => x.Quantity)
    .Column(x => x.UnitPrice)
    .Column(x => x.Total, column => column
        .Formula(context => $"=A{context.Row}*B{context.Row}")
        .Format("#,##0.00")
        .Align(ExcelHorizontalAlignment.Right))
    .Build();
```

Numeric format codes are locale-invariant. For example, `#,##0.00` displays `1` as `1,00` in Excel with an Italian locale; use `.` for the decimal placeholder.

The callback also exposes the one-based output `Column` and `SheetName`. Formula text may include one leading `=` or omit it; CellSharp writes the normalized XLSX expression without that prefix. Empty formulas and a double `==` prefix fail during export.

`Formula(...)` is an explicit executable-content boundary. Do not concatenate user-controlled values, imported workbook text, or other untrusted data into the returned expression: spreadsheet applications may evaluate network-capable or otherwise dangerous functions. Write untrusted data through a normal column and reference that cell from a constant, application-authored formula instead.

For a formula column, the formula remains the value Excel displays and recalculates. CellSharp reads the property value and applies its converter's `Write` method only to serialize the initial cached result. Keep that property value consistent with the formula if the workbook is consumed before Excel or LibreOffice recalculates it. The column's format, width, and alignment still apply. Normal string values remain literal inline text even if they begin with `=`.

Formulas work for single-sheet and multi-sheet writes. Cross-sheet references are passed through unchanged, for example `=Source!A2`. A generated template always remains header-only and contains no formula cells.

## Recalculation and import

When an export writes at least one formula data row, CellSharp requests full workbook recalculation on load and save. The exported cache makes immediate import possible, but CellSharp cannot calculate a formula itself: it is the caller's responsibility to provide an initial property value consistent with the formula.

On import, a formula cell's cached value uses the same scalar conversion, `ConvertWith`, and validation flow as a normal cell. A formula with no cached value is an `InvalidValue` error. The default invalid-row policy skips that row; `ExcelInvalidRowPolicy.Include` retains its partially populated object.

Formula columns cannot use `AllowedValues`, `Range`, or `DateBetween`, because those produce native Excel Data Validation for editable input cells. Custom `.Validate(...)` remains an import-time rule, and `Optional()` keeps its usual header-mapping meaning.

For a small executable order-total example, see [Formulas and calculated columns](use-cases/formulas-and-calculated-columns.md).

# Named ranges in formulas

Workbook-scoped names created with `Range(...).Name(...)` are ordinary Excel names. Use them directly in a formula; CellSharp does not rewrite the expression.

```csharp
sheet.Range("D2:D100").Name("Revenue");
sheet.Cell("D101").Formula("SUM(Revenue)");
```
