# Streams and web APIs

## Scenario

An HTTP endpoint needs to produce an XLSX response without creating a temporary server file.

## Write

```csharp
app.MapGet("/customers.xlsx", () =>
{
    var customers = customerService.GetAll();
    using var stream = new MemoryStream();

    Excel.Write(stream, customers);
    return Results.File(
        stream.ToArray(),
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "customers.xlsx");
});
```

## Read

```csharp
using var stream = new MemoryStream(uploadedBytes);
var result = Excel.Read<Customer>(stream);
```

## Result

CellSharp reads and writes seekable streams without taking ownership. Writes reset to position zero, truncate the stream, and leave it open at the end. Reads start from position zero and also leave the stream open.

## Why this approach

The library has no web-framework dependency; the endpoint only adapts its application’s bytes and headers around a normal CellSharp stream call.

## Try it

Run [Streams](../../samples/CellSharp.Samples/Scenarios/Streams.cs). For cancellation and multi-sheet streams, see [streams](../streams.md).
