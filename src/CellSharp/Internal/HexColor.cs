namespace CellSharp.Internal;

internal static class HexColor
{
    internal static string Normalize(string color, string parameterName)
    {
        if (color is null)
        {
            throw new ArgumentNullException(parameterName);
        }

        if (color.Length != 7 || color[0] != '#' || !color.Skip(1).All(IsHex))
        {
            throw new ArgumentException("Colors must use the #RRGGBB format.", parameterName);
        }

        return color.Substring(1).ToUpperInvariant();
    }

    private static bool IsHex(char value) => (value >= '0' && value <= '9')
        || (value >= 'A' && value <= 'F')
        || (value >= 'a' && value <= 'f');
}
