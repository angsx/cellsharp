using System.Text.RegularExpressions;

namespace CellSharp.Internal;

internal static class ExcelDefinedName
{
    private const int MaximumLength = 255;
    private static readonly Regex A1Reference = new("^[A-Za-z]{1,3}[1-9][0-9]*$", RegexOptions.CultureInvariant);
    private static readonly Regex R1C1Reference = new("^R[0-9]+C[0-9]+$", RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);
    private static readonly HashSet<string> ReservedNames = new(StringComparer.OrdinalIgnoreCase) { "R", "C", "Print_Area", "Print_Titles", "_FilterDatabase" };

    internal static void Validate(string name, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("A defined name is required.", parameterName);
        if (name.Length > MaximumLength) throw new ArgumentException($"A defined name cannot exceed {MaximumLength} characters.", parameterName);
        if (!IsFirst(name[0]) || name.Any(character => !IsPart(character))) throw new ArgumentException("A defined name must start with a letter, underscore, or backslash and contain only letters, digits, underscores, periods, or backslashes.", parameterName);
        if (A1Reference.IsMatch(name) || R1C1Reference.IsMatch(name) || ReservedNames.Contains(name) || name.StartsWith("_xlnm.", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException($"'{name}' is not a valid workbook defined name.", parameterName);
    }

    internal static string AbsoluteReference(string sheetName, ExcelRangeReference range)
    {
        var escapedSheetName = sheetName.Replace("'", "''");
        if (range.FromRow == 1 && range.ToRow == ExcelRangeReference.MaximumRows)
            return $"'{escapedSheetName}'!${ExcelRangeReference.ColumnName(range.FromColumn)}:${ExcelRangeReference.ColumnName(range.ToColumn)}";
        var start = $"${ExcelRangeReference.ColumnName(range.FromColumn)}${range.FromRow}";
        var end = $"${ExcelRangeReference.ColumnName(range.ToColumn)}${range.ToRow}";
        return $"'{escapedSheetName}'!{(range.IsCell ? start : start + ":" + end)}";
    }

    private static bool IsFirst(char value) => char.IsLetter(value) || value is '_' or '\\';
    private static bool IsPart(char value) => char.IsLetterOrDigit(value) || value is '_' or '.' or '\\';
}

internal static class HeaderFooterText
{
    private static readonly HashSet<char> Tokens = new() { 'P', 'N', 'D', 'T', 'F', 'A', '&' };
    internal static string Normalize(string value, string parameterName)
    {
        if (value is null) throw new ArgumentNullException(parameterName);
        var result = new System.Text.StringBuilder(value.Length);
        for (var index = 0; index < value.Length; index++)
        {
            var character = value[index];
            if (character == '&' && (index + 1 == value.Length || !Tokens.Contains(char.ToUpperInvariant(value[index + 1])))) result.Append("&&");
            else result.Append(character);
        }
        return result.ToString();
    }
}
