using CellSharp.Internal;

namespace CellSharp;

/// <summary>Configures the left, center, and right sections of a worksheet header or footer.</summary>
public sealed class ExcelHeaderFooter
{
    private readonly HeaderFooterDefinition _definition;
    internal ExcelHeaderFooter(HeaderFooterDefinition definition) => _definition = definition;

    /// <summary>Sets the left section. Excel tokens such as &amp;P and &amp;N are preserved.</summary>
    public ExcelHeaderFooter Left(string value) { _definition.SetLeft(HeaderFooterText.Normalize(value, nameof(value))); return this; }
    /// <summary>Sets the center section. Excel tokens such as &amp;D and &amp;A are preserved.</summary>
    public ExcelHeaderFooter Center(string value) { _definition.SetCenter(HeaderFooterText.Normalize(value, nameof(value))); return this; }
    /// <summary>Sets the right section. Excel tokens such as &amp;F and &amp;T are preserved.</summary>
    public ExcelHeaderFooter Right(string value) { _definition.SetRight(HeaderFooterText.Normalize(value, nameof(value))); return this; }
}
