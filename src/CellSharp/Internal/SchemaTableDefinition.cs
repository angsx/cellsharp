using System.Text.RegularExpressions;

namespace CellSharp.Internal;

internal sealed class SchemaTableDefinition
{
    private static readonly Regex ValidName = new("^[A-Za-z_][A-Za-z0-9_.]{0,254}$", RegexOptions.CultureInvariant);
    private static readonly Regex CellReference = new("^[A-Za-z]{1,3}[0-9]+$", RegexOptions.CultureInvariant);
    private static readonly Regex R1C1Reference = new("^R[0-9]+C[0-9]+$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    internal SchemaTableDefinition(string? name, string? style)
    {
        if (name is not null)
        {
            Name = ValidateName(name, nameof(name));
        }

        if (style is not null && (string.IsNullOrWhiteSpace(style) || !style.StartsWith("TableStyle", StringComparison.Ordinal)))
        {
            throw new ArgumentException("A table style must be a native Excel TableStyle name or null.", nameof(style));
        }

        Style = style;
    }

    internal string? Name { get; }

    internal string? Style { get; }

    internal static string ValidateName(string name, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(name))
        {
            throw new ArgumentException("A table name is required.", parameterName);
        }

        if (!ValidName.IsMatch(name) || CellReference.IsMatch(name) || R1C1Reference.IsMatch(name)
            || string.Equals(name, "R", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "C", StringComparison.OrdinalIgnoreCase))
        {
            throw new ArgumentException("The table name must be an Excel-compatible identifier and cannot be a cell reference.", parameterName);
        }

        return name;
    }
}
