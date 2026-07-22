namespace CellSharp.Internal;

internal sealed class HeaderStyleOverride
{
    internal HeaderStyleOverride(bool? bold, string? background, string? foreground)
    {
        Bold = bold;
        Background = background;
        Foreground = foreground;
    }

    internal bool? Bold { get; }

    internal string? Background { get; }

    internal string? Foreground { get; }
}
