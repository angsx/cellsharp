namespace CellSharp.Internal;

internal static class ExcelSheetName
{
    internal const string Default = "Sheet1";

    internal static string Validate(string sheetName, string parameterName)
    {
        if (string.IsNullOrWhiteSpace(sheetName))
        {
            throw new ArgumentException("A worksheet name is required.", parameterName);
        }

        if (sheetName.Length > 31)
        {
            throw new ArgumentOutOfRangeException(
                parameterName,
                "An Excel worksheet name cannot exceed 31 characters.");
        }

        if (sheetName.Any(character => char.IsControl(character)
            || character is ':' or '\\' or '/' or '?' or '*' or '[' or ']'))
        {
            throw new ArgumentException(
                "An Excel worksheet name cannot contain control characters or any of : \\ / ? * [ ].",
                parameterName);
        }

        return sheetName;
    }
}
