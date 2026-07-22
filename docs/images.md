# Images

Add native PNG or JPEG worksheet images without handling drawing parts or EMU units. Positions, offsets, and dimensions use pixels.

```csharp
sheet.AddImage("logo.png")
    .At("A1");
```

Images use a one-cell anchor: they remain attached to their top-left cell when worksheet row heights or column widths change.

## Resize and aspect ratio

An image with no explicit size uses its natural pixel dimensions. `Size` is always explicit; for one dimension, use `KeepAspectRatio` to calculate the other dimension from the original image.

```csharp
sheet.AddImage("logo.png")
    .At("A1", offsetX: 8, offsetY: 6)
    .Width(160)
    .KeepAspectRatio();

sheet.AddImage("badge.jpg")
    .At(row: 3, column: 2)
    .Size(240, 120);
```

`Size(width, height)` takes precedence over aspect-ratio preservation. Configuring both `Width` and `Height` together with `KeepAspectRatio` is rejected because the requested result would be ambiguous.

## Streams

Only PNG and JPEG are supported. For streams, specify the format explicitly.

| Format | Path extension | Stream enum |
| --- | --- | --- |
| PNG | `.png` | `ExcelImageFormat.Png` |
| JPEG | `.jpg`, `.jpeg` | `ExcelImageFormat.Jpeg` |

```csharp
using var stream = File.OpenRead("logo.png");

sheet.AddImage(stream, ExcelImageFormat.Png)
    .At("B2");
```

The caller owns the stream. CellSharp reads and snapshots its bytes immediately during `AddImage`, does not require seeking, and never closes the supplied stream. PNG/JPEG signatures and dimensions are validated without `System.Drawing`.

## Layout example

Image coordinates are absolute worksheet coordinates, so a logo naturally stays above a dataset offset with `DataStartAt`.

```csharp
Excel.Workbook().AddSheet(rows, schema, sheet =>
{
    sheet.DataStartAt("A7");
    sheet.AddImage("logo.png").At("A1").Width(140).KeepAspectRatio();
    sheet.Merge("B1:F2").Value("Sales report").Style(s => s.Bold().FontSize(18));
    sheet.Footer.Right("Page &P of &N");
});
```

Images coexist with tables, validations, conditional formatting, comments, merged ranges, headers/footers, and page breaks.

## Remote demo asset

`Images.xlsx` in the samples project downloads one PNG placeholder over HTTPS from a fixed allowlisted host before passing its bytes to `AddImage(stream, ExcelImageFormat.Png)`. This demonstrates the caller-owned stream API; CellSharp itself accepts only file paths and streams, never image URLs. Applications should apply their own network policy, host allowlist, timeout, and size limits before supplying downloaded bytes.

## Limitations

This first version intentionally does not support image hyperlinks, crop, rotation, opacity/effects, shapes, charts, arbitrary anchor modes, or cross-sheet media deduplication.
