namespace CellSharp.Internal;

internal static class ReportComponentDefaults
{
    internal static void Title(ExcelRange range) => range.Style(style => style.Bold().FontSize(20).AlignCenter().VerticalAlignCenter().WrapText());
    internal static void Section(ExcelRange range) => range.Style(style => style.Bold().FontSize(12).VerticalAlignCenter().FillColor("#F2F2F2").Border(border => border.Bottom(ExcelBorderStyle.Thin, "#D9D9D9")));
    internal static void Note(ExcelRange range) => range.Style(style => style.Italic().FontSize(10).FontColor("#666666").VerticalAlignTop().WrapText());
    internal static void KpiContainer(ExcelRange range) => range.Style(style => style.FillColor("#FAFAFA").Border(border => border.Outline(ExcelBorderStyle.Thin, "#D9D9D9")));
    internal static void KpiLabel(ExcelRange range) => range.Style(style => style.Bold().FontSize(10).VerticalAlignBottom());
    internal static void KpiValue(ExcelRange range) => range.Style(style => style.Bold().FontSize(18).VerticalAlignCenter());
}
