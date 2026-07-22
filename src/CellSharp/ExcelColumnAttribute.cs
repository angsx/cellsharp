namespace CellSharp;

/// <summary>Declares stable worksheet metadata for a property used by <see cref="Excel.SchemaFromAttributes{T}()"/>.</summary>
/// <remarks>
/// Properties without this attribute are still included by convention. Assigning <see cref="Order"/>
/// makes a column ordered explicitly; explicitly ordered columns precede unordered columns.
/// </remarks>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ExcelColumnAttribute : Attribute
{
    private int _order = int.MaxValue;

    /// <summary>Initializes an attribute without overriding the property name used as the header.</summary>
    public ExcelColumnAttribute()
    {
    }

    /// <summary>Initializes an attribute with the worksheet header for the property.</summary>
    public ExcelColumnAttribute(string header)
    {
        Header = header;
    }

    /// <summary>Gets the worksheet header, or <see langword="null"/> to use the property name.</summary>
    public string? Header { get; }

    /// <summary>
    /// Gets or sets the explicit column order. If it is not assigned, the property retains normal
    /// reflection discovery order after all explicitly ordered columns.
    /// </summary>
    public int Order
    {
        get => _order;
        set
        {
            _order = value;
            HasOrder = true;
        }
    }

    /// <summary>Gets or sets whether this column's input header may be absent when reading.</summary>
    public bool Optional { get; set; }

    /// <summary>Gets or sets the Excel number or date format code used when exporting this column.</summary>
    public string? Format { get; set; }

    internal bool HasOrder { get; private set; }
}
