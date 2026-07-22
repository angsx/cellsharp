# Streams and cancellation

Every public read, write, template, multi-sheet, and `Excel.Open(...)` workflow has seekable-stream support where the operation needs it. The caller owns the stream: CellSharp validates capabilities but never disposes it.

```csharp
using var stream = new MemoryStream();

Excel.Write(stream, orders, schema);
stream.Position = 0;
var result = Excel.Read(stream, schema);
```

Writes reset the stream to position zero, truncate it, and finish at `Position == Length`. Reads begin at position zero. Streams must expose the required readable/writable and seekable capabilities; a non-seekable stream is rejected before CellSharp opens an XLSX package.

Selected synchronous stream operations accept a final `CancellationToken`:

```csharp
Excel.Write(stream, orders, schema, cancellationToken);
```

Cancellation is cooperative and is observed before work and between rows or worksheets. It throws `OperationCanceledException`, not an import diagnostic. A cancelled write can leave partial content in the caller-owned stream. CellSharp intentionally has no `Async` wrappers because OpenXML package I/O is synchronous.

See [async and cancellation](async-and-cancellation.md) for the design rationale.

For an ASP.NET-style endpoint without a library web dependency, see [Streams and web APIs](use-cases/streams-and-web-apis.md).
