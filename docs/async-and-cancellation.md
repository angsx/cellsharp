# Async and cancellation

CellSharp 0.7 does not add `WriteAsync`, `ReadAsync`, or `OpenAsync`. DocumentFormat.OpenXml 3.3.0 exposes synchronous package create/open/save operations; wrapping them in `Task.Run` would consume a worker thread while presenting misleading async I/O semantics.

Instead, selected synchronous APIs accept a final `CancellationToken` and check it cooperatively before work and between worksheet rows, sheets, and native validation lookup work:

```csharp
using var stream = new MemoryStream();
using var cancellation = new CancellationTokenSource();

Excel.Write(stream, rows, schema, cancellation.Token);
```

Cancellation throws `OperationCanceledException`. It is not converted into `ExcelReadError`, regardless of `ExcelInvalidRowPolicy`, and exceptions of that type raised by converters or validators also propagate.

The checks do not run for every cell. A cancellation requested while one row is being converted is observed before the next row, which keeps normal per-cell overhead unchanged.

Stream ownership from 0.6 remains unchanged: caller-owned streams remain open after success or cancellation. A cancelled write can leave a partial or invalid workbook in the target stream; CellSharp does not promise atomic writes and does not create hidden temporary copies. Reads, writes, and templates remain synchronous and use seekable streams directly.
