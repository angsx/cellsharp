namespace CellSharp;

/// <summary>Converts one logical property value to and from a CellSharp-supported cell value.</summary>
/// <typeparam name="TValue">The logical property type.</typeparam>
/// <typeparam name="TCellValue">The scalar representation stored in the XLSX cell.</typeparam>
public interface IExcelValueConverter<TValue, TCellValue>
{
    /// <summary>Converts a logical value to its scalar XLSX representation.</summary>
    TCellValue Write(TValue value);

    /// <summary>Attempts to convert a scalar XLSX representation to its logical value.</summary>
    bool TryRead(TCellValue value, out TValue converted);
}
