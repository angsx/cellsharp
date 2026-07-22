namespace CellSharp.Internal;

internal abstract class ValueConverterDefinition
{
    internal abstract Type CellType { get; }

    internal abstract object? Write(object value);

    internal abstract bool TryRead(object value, out object? converted);
}

internal sealed class ValueConverterDefinition<TValue, TCellValue> : ValueConverterDefinition
{
    private readonly IExcelValueConverter<TValue, TCellValue> _converter;

    internal ValueConverterDefinition(IExcelValueConverter<TValue, TCellValue> converter)
    {
        _converter = converter;
    }

    internal override Type CellType => typeof(TCellValue);

    internal override object? Write(object value) => _converter.Write((TValue)value);

    internal override bool TryRead(object value, out object? converted)
    {
        if (_converter.TryRead((TCellValue)value, out var typed))
        {
            converted = typed;
            return true;
        }

        converted = null;
        return false;
    }
}
