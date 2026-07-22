# Excel toolkit milestone plan

CellSharp keeps four deliberate layers: typed data, worksheet primitives, Excel-native capabilities, and convenience report components. A report helper must be reducible to the worksheet primitives; it must not require a separate writer pipeline.

## Completed foundation

- Audit: the layout writer runs after schema cells, formula/table/validation offsets are based on `DataStartAt`, and workbook recalculation is requested whenever formulas are emitted.
- Property-level style composition: theme/schema, column, row, range, and cell styles now compose in that order. Booleans are nullable internally, so an explicit `false` remains distinguishable from an unspecified property; border sides compose independently.
- Resolved-style deduplication: equivalent style chains share one cell format, font, fill, and border record.
- Selective AutoFit: `Column(...).AutoFit()`, `Columns(...).AutoFit()`, and `Range("F:H").AutoFitColumns()` examine only materialized cells. Explicit widths win.
- Conditional formatting: native comparison, formula, text, duplicate/unique, and blank rules use deterministic priorities and deduplicated differential styles. See [conditional formatting](conditional-formatting.md).
- Workbook/worksheet utilities: workbook-scoped [named ranges](named-ranges.md), classic Excel comments, printable headers/footers, and manual row/column page breaks.
- Images: native PNG/JPEG drawings with cell anchors, pixel sizing, stream snapshots, and no graphics-library dependency. See [images](images.md).
- Report conveniences: `Title`, `Section`, `Note`, and `Kpi` expand into ordinary values, merged ranges, and composable styles. See [report layouts](report-layouts.md).

## Milestone status

The completed slices include Open XML validation, relationship checks where relevant, existing-suite regression coverage, and real sample workbooks. Charts, rich text, themes, dashboards, and structural row/column insertion remain intentionally outside this milestone.

## Near-term

- Additional conditional-formatting rule families only after an API design review.
- Image refinements that preserve the current PNG/JPEG, cell-anchored model.
- A report theme only if a small, stable customization model emerges from real use.

## Major future milestone

- Charts.

## Later

- Shapes, sparklines, richer drawings, and pivot-oriented features.

## Compatibility and limits

The basic APIs (`Excel.Write`, `Excel.Read`, `Excel.CreateTemplate`, `Excel.Workbook`, and `Excel.Open`) remain unchanged. XLSX coordinate validation is centralized in the worksheet reference type, including the 1,048,576-row and 16,384-column (`XFD`) limits.
