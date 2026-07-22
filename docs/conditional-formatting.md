# Conditional formatting

Apply native Excel conditional formatting directly to a cell or range. Rules do not create or style cells eagerly, so a whole-column range remains sparse.

```csharp
sheet.Range("D2:D100")
    .ConditionalFormat()
    .GreaterThan(1000)
    .Style(s => s.FillColor("#C6EFCE").FontColor("#006100"));
```

## Comparisons

Use `GreaterThan`, `GreaterThanOrEqual`, `LessThan`, `LessThanOrEqual`, `EqualTo`, `NotEqualTo`, `Between`, and `NotBetween`. Numeric values, `DateTime`, booleans, and text values are normalized for Excel using invariant conventions.

```csharp
sheet.Range("C2:C500").ConditionalFormat().LessThan(0)
    .Style(s => s.FillColor("#FFC7CE").FontColor("#9C0006"))
    .StopIfTrue();

sheet.Range("D2:D500").ConditionalFormat().Between(0.20m, 0.40m)
    .Style(s => s.FillColor("#FFEB9C"));
```

Rules receive deterministic, one-based priorities in declaration order. `StopIfTrue()` prevents later matching rules from being applied by Excel.

## Formula and text rules

Formula rules accept a leading `=` but do not require one. The formula is evaluated relative to the top-left cell of the conditional range, as in Excel itself.

```csharp
sheet.Range("E2:E100")
    .ConditionalFormat()
    .Formula("=E2<TODAY()")
    .Style(s => s.FontColor("#9C0006"));

sheet.Range("F2:F100").ConditionalFormat().ContainsText("ERROR")
    .Style(s => s.Bold().FillColor("#FFC7CE"));
sheet.Range("F2:F100").ConditionalFormat().BeginsWith("WARN");
sheet.Range("F2:F100").ConditionalFormat().EndsWith("!");
```

`DuplicateValues`, `UniqueValues`, `Blanks`, and `NonBlanks` are also native Excel rules.

## Differential styles

Conditional formatting uses differential styles (DXF), not normal cell styles. `Style(...)` supports font decorations and color, fill color, borders, and number formats. Equivalent differential styles are shared in the workbook. Font name/size, alignment, rotation, wrap, shrink-to-fit, and inside borders are rejected because they are not supported by this stable DXF surface.

## Interaction with typed sheets

Ranges always use absolute worksheet coordinates. `DataStartAt` changes where data is written; it does not change how a conditional-formatting reference is interpreted. Conditional formatting is independent from tables, validations, and merges.

```csharp
Excel.Workbook().AddSheet(rows, schema, sheet =>
{
    sheet.DataStartAt("B4");
    sheet.Range("D5:D500").ConditionalFormat().LessThan(0)
        .Style(s => s.FillColor("#FFC7CE"));
});
```

Color scales, data bars, icon sets, and top/bottom or average rules are planned follow-up work.
