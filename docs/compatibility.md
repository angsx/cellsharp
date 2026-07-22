# Compatibility

CellSharp targets `netstandard2.0` and `net8.0`.

`netstandard2.0` keeps the core library usable by a broad range of .NET implementations. `net8.0` provides a current runtime target without changing the public API. The same CellSharp API is available from both targets.

## Supported scalar values

- `string`
- `byte`, `sbyte`, `short`, `ushort`, `int`, `uint`, `long`, `ulong`, `float`, `double`, and `decimal`
- `bool`
- `DateTime`
- `Guid`
- Nullable value-type variants of the above

Writing requires public readable scalar properties. Reading requires a public parameterless constructor and public writable scalar properties. A typed schema requires configured properties to be both public readable and writable.

## Deliberately out of scope

CellSharp provides built-in themes, immutable visual templates, header customization, alternating rows, typed multi-sheet export/import/template generation, runtime schema overlays, runtime import mapping, per-column format/width/alignment, declarative native validation, converters, formulas, Excel Tables, worksheet filters/panes, and essential print settings. It deliberately does not provide conditional formatting, charts, pivot tables, macros, images, comments, hyperlinks, table append/resize or totals rows, a structured-reference builder, advanced print layout, `DataTable`, JSON, public streaming primitives, or header-localization logic.

See [writing](writing.md), [reading](reading.md), [schemas](schemas.md), and [styling](styling.md) for the behavior of the supported API.
