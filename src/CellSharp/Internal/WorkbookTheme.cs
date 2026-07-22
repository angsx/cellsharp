namespace CellSharp.Internal;

internal static class WorkbookTheme
{
    internal static ExcelStyleTemplate For(ExcelTheme theme)
    {
        return theme switch
        {
            ExcelTheme.Default => new ExcelStyleTemplate(
                headerTextColor: "#1F2937", headerBackgroundColor: "#F3F4F6", borderColor: "#D1D5DB"),
            ExcelTheme.Modern => new ExcelStyleTemplate(
                fontName: "Aptos", dataTextColor: "#1F2937", dataBackgroundColor: "#F8FAFC",
                headerTextColor: "#FFFFFF", headerBackgroundColor: "#1F4E78",
                alternateRowBackgroundColor: "#EDF4FC", borderColor: "#D9E2F3"),
            ExcelTheme.Classic => new ExcelStyleTemplate(
                dataTextColor: "#000000", dataBackgroundColor: "#FFFFFF",
                headerTextColor: "#274E13", headerBackgroundColor: "#D9EAD3",
                alternateRowBackgroundColor: "#F3F8F0", borderColor: "#93C47D"),
            ExcelTheme.Minimal => new ExcelStyleTemplate(
                fontName: "Arial", fontSize: 10D, dataTextColor: "#1F2937", dataBackgroundColor: "#FFFFFF",
                headerTextColor: "#1F2937", headerBackgroundColor: "#FFFFFF",
                alternateRowBackgroundColor: "#F9FAFB", borderColor: "#E5E7EB"),
            _ => throw new ArgumentOutOfRangeException(nameof(theme)),
        };
    }
}
