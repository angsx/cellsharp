using CellSharp.Internal;

namespace CellSharp;

/// <summary>Defines an immutable, reusable visual identity for XLSX exports.</summary>
public sealed class ExcelStyleTemplate
{
    /// <summary>Creates a visual template. Color values must use the #RRGGBB format.</summary>
    public ExcelStyleTemplate(
        string fontName = "Calibri",
        double fontSize = 11D,
        string dataTextColor = "#111827",
        string dataBackgroundColor = "#FFFFFF",
        string? headerFontName = null,
        double? headerFontSize = null,
        string headerTextColor = "#1F2937",
        string headerBackgroundColor = "#F3F4F6",
        bool headerBold = true,
        string? alternateRowBackgroundColor = "#F8FAFC",
        string borderColor = "#E5E7EB")
    {
        FontName = RequiredName(fontName, nameof(fontName));
        FontSize = FontSizeValue(fontSize, nameof(fontSize));
        DataTextColor = Color(dataTextColor, nameof(dataTextColor));
        DataBackgroundColor = Color(dataBackgroundColor, nameof(dataBackgroundColor));
        HeaderFontName = RequiredName(headerFontName ?? fontName, nameof(headerFontName));
        HeaderFontSize = FontSizeValue(headerFontSize ?? fontSize, nameof(headerFontSize));
        HeaderTextColor = Color(headerTextColor, nameof(headerTextColor));
        HeaderBackgroundColor = Color(headerBackgroundColor, nameof(headerBackgroundColor));
        HeaderBold = headerBold;
        AlternateRowBackgroundColor = alternateRowBackgroundColor is null
            ? null
            : Color(alternateRowBackgroundColor, nameof(alternateRowBackgroundColor));
        BorderColor = Color(borderColor, nameof(borderColor));
    }

    /// <summary>Gets the data-cell font family.</summary>
    public string FontName { get; }

    /// <summary>Gets the data-cell font size in points.</summary>
    public double FontSize { get; }

    /// <summary>Gets the data-cell text color in #RRGGBB form.</summary>
    public string DataTextColor { get; }

    /// <summary>Gets the data-cell background color in #RRGGBB form.</summary>
    public string DataBackgroundColor { get; }

    /// <summary>Gets the header font family.</summary>
    public string HeaderFontName { get; }

    /// <summary>Gets the header font size in points.</summary>
    public double HeaderFontSize { get; }

    /// <summary>Gets the header text color in #RRGGBB form.</summary>
    public string HeaderTextColor { get; }

    /// <summary>Gets the header background color in #RRGGBB form.</summary>
    public string HeaderBackgroundColor { get; }

    /// <summary>Gets whether header text is bold.</summary>
    public bool HeaderBold { get; }

    /// <summary>Gets the alternate data-row background, or null to use the data background.</summary>
    public string? AlternateRowBackgroundColor { get; }

    /// <summary>Gets the cell border color in #RRGGBB form.</summary>
    public string BorderColor { get; }

    private static string Color(string value, string parameterName) => $"#{HexColor.Normalize(value, parameterName)}";

    private static string RequiredName(string value, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException("A font name is required.", parameterName);
        }

        return value;
    }

    private static double FontSizeValue(double value, string parameterName)
    {
        if (value <= 0D || value > 409D)
        {
            throw new ArgumentOutOfRangeException(parameterName, "Font size must be greater than zero and no greater than 409.");
        }

        return value;
    }
}
