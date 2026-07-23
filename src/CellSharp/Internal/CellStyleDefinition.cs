namespace CellSharp.Internal;

internal sealed class CellStyleDefinition : IEquatable<CellStyleDefinition>
{
    internal CellStyleDefinition(
        string fontName,
        double fontSize,
        bool bold,
        string foreground,
        string background,
        string border,
        ExcelHorizontalAlignment alignment,
        ExcelVerticalAlignment verticalAlignment)
    {
        FontName = fontName;
        FontSize = fontSize;
        Bold = bold;
        Foreground = foreground;
        Background = background;
        Border = border;
        Alignment = alignment;
        VerticalAlignment = verticalAlignment;
    }

    internal string FontName { get; }

    internal double FontSize { get; }

    internal bool Bold { get; }

    internal string Foreground { get; }

    internal string Background { get; }

    internal string Border { get; }

    internal ExcelHorizontalAlignment Alignment { get; }

    internal ExcelVerticalAlignment VerticalAlignment { get; }

    internal CellStyleDefinition With(HeaderStyleOverride? overrideStyle)
    {
        if (overrideStyle is null)
        {
            return this;
        }

        return new CellStyleDefinition(
            FontName,
            FontSize,
            overrideStyle.Bold ?? Bold,
            overrideStyle.Foreground ?? Foreground,
            overrideStyle.Background ?? Background,
            Border,
            Alignment,
            VerticalAlignment);
    }

    internal CellStyleDefinition WithBackground(string background) => new(
        FontName,
        FontSize,
        Bold,
        Foreground,
        background,
        Border,
        Alignment,
        VerticalAlignment);

    internal CellStyleDefinition WithAlignment(ExcelHorizontalAlignment? alignment, ExcelVerticalAlignment? verticalAlignment) => alignment is null && verticalAlignment is null
        ? this
        : new CellStyleDefinition(
            FontName,
            FontSize,
            Bold,
            Foreground,
            Background,
            Border,
            alignment ?? Alignment,
            verticalAlignment ?? VerticalAlignment);

    public bool Equals(CellStyleDefinition? other)
    {
        return other is not null
            && FontName == other.FontName
            && FontSize.Equals(other.FontSize)
            && Bold == other.Bold
            && Foreground == other.Foreground
            && Background == other.Background
            && Border == other.Border
            && Alignment == other.Alignment
            && VerticalAlignment == other.VerticalAlignment;
    }

    public override bool Equals(object? obj) => Equals(obj as CellStyleDefinition);

    public override int GetHashCode()
    {
        unchecked
        {
            var hash = 17;
            hash = (hash * 31) + FontName.GetHashCode();
            hash = (hash * 31) + FontSize.GetHashCode();
            hash = (hash * 31) + Bold.GetHashCode();
            hash = (hash * 31) + Foreground.GetHashCode();
            hash = (hash * 31) + Background.GetHashCode();
            hash = (hash * 31) + Border.GetHashCode();
            hash = (hash * 31) + Alignment.GetHashCode();
            hash = (hash * 31) + VerticalAlignment.GetHashCode();
            return hash;
        }
    }
}
