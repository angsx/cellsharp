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
        ExcelHorizontalAlignment alignment)
    {
        FontName = fontName;
        FontSize = fontSize;
        Bold = bold;
        Foreground = foreground;
        Background = background;
        Border = border;
        Alignment = alignment;
    }

    internal string FontName { get; }

    internal double FontSize { get; }

    internal bool Bold { get; }

    internal string Foreground { get; }

    internal string Background { get; }

    internal string Border { get; }

    internal ExcelHorizontalAlignment Alignment { get; }

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
            Alignment);
    }

    internal CellStyleDefinition WithBackground(string background) => new(
        FontName,
        FontSize,
        Bold,
        Foreground,
        background,
        Border,
        Alignment);

    internal CellStyleDefinition WithAlignment(ExcelHorizontalAlignment? alignment) => alignment is null
        ? this
        : new CellStyleDefinition(
            FontName,
            FontSize,
            Bold,
            Foreground,
            Background,
            Border,
            alignment.Value);

    public bool Equals(CellStyleDefinition? other)
    {
        return other is not null
            && FontName == other.FontName
            && FontSize.Equals(other.FontSize)
            && Bold == other.Bold
            && Foreground == other.Foreground
            && Background == other.Background
            && Border == other.Border
            && Alignment == other.Alignment;
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
            return hash;
        }
    }
}
