# Custom converters

Use a converter when a property is a domain type, a legacy code, or another representation that CellSharp does not support directly. A converter works on logical values and CellSharp scalar cell values; it never sees Open XML objects, workbook state, or styles.

## Reusable converter

Implement `IExcelValueConverter<TValue, TCellValue>`, where `TCellValue` is a CellSharp-supported scalar such as `string`, a numeric type, `bool`, `DateTime`, or `Guid`.

```csharp
public sealed class CustomerStatusConverter
    : IExcelValueConverter<CustomerStatus?, string>
{
    public string Write(CustomerStatus? value) => value?.Code ?? string.Empty;

    public bool TryRead(string value, out CustomerStatus? converted)
    {
        converted = CustomerStatus.FromCode(value);
        return converted is not null;
    }
}

public static readonly CustomerStatusConverter StatusConverter = new();
```

Attach the same stateless instance to a schema column:

```csharp
.Column(x => x.Status, column => column.ConvertWith(StatusConverter))
```

## Read and write behavior

On export CellSharp calls `Write`, then writes the returned scalar through its normal XLSX writer. On import CellSharp first converts the raw cell to `TCellValue`, then calls `TryRead`. This scalar step uses the invariant XLSX rules followed by any fallback `ExcelReadOptions.Culture` for numeric and `DateTime` text values. A `false` return becomes an `InvalidValue` `ExcelReadError`; other valid rows are retained. Exceptions from a converter propagate because they represent converter bugs or configuration failures rather than expected bad workbook data.

## Nullability

Blank cells bypass the converter. A nullable property receives `null`; non-nullable properties retain CellSharp's existing blank-cell conversion behavior. An explicit empty text cell is a present value by default. With `ExcelReadOptions.EmptyStringAsNull` enabled, it is treated like a blank cell and also bypasses the converter. Converters should therefore handle values that are present, not use `null` to implement requiredness.

## Format and validation

Converters and `Format(...)` are separate. For example, a `Money` converter may return `decimal`, after which `Format("#,##0.00")` controls only Excel presentation.

The runtime order is:

```text
raw cell → scalar conversion → custom converter → declarative validation → custom Validate
```

Validators receive `TValue`, not the scalar representation. Declarative native Excel validation is only meaningful when the configured rule is compatible with the written scalar representation; CellSharp does not infer or translate domain values automatically. See [validation](validation.md).

## Unsupported property types

Without a converter, a schema column still requires a CellSharp-supported scalar type. With `ConvertWith`, a non-native property type is valid when its declared `TCellValue` is supported. Converters are metadata on a schema column, so they also work with `Excel.CreateTemplate`; templates retain format, width, alignment, and any independently compatible native validation.

For a complete write-and-read example, see [Custom converters](use-cases/custom-converters.md) and run its linked sample.
