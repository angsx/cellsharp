# Named ranges

Declare a workbook-scoped Excel name from any worksheet range. CellSharp writes an absolute, sheet-qualified reference, so formulas can use the name naturally.

```csharp
sheet.Range("D2:D100")
    .Name("Revenue");

sheet.Cell("D101")
    .Formula("SUM(Revenue)");
```

Names use absolute worksheet coordinates; `DataStartAt` does not make them dataset-relative. Single cells, rectangular ranges, and whole-column ranges are supported.

```csharp
sheet.Cell("B2").Name("CustomerId");
sheet.Range("C:C").Name("AllCosts");
```

Names are workbook-scoped in this release. They must be unique case-insensitively, be at most 255 characters, start with a letter, underscore, or backslash, and cannot resemble an A1/R1C1 reference or use reserved Excel name forms. Worksheet-local names, dynamic ranges, and named formulas are intentionally deferred.
