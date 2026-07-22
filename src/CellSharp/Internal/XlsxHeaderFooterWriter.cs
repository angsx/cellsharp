using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxHeaderFooterWriter
{
    internal static void Apply(IReadOnlyList<WorksheetValidationContext> worksheets)
    {
        foreach (var context in worksheets)
        {
            var layout = context.Layout;
            if (layout is null || (!layout.Header.HasValue && !layout.Footer.HasValue)) continue;
            var headerFooter = new HeaderFooter();
            if (layout.Header.HasValue) headerFooter.OddHeader = new OddHeader(Text(layout.Header));
            if (layout.Footer.HasValue) headerFooter.OddFooter = new OddFooter(Text(layout.Footer));
            context.Worksheet.AppendChild(headerFooter);
        }
    }

    private static string Text(HeaderFooterDefinition definition) =>
        (definition.Left is null ? string.Empty : "&L" + definition.Left) +
        (definition.Center is null ? string.Empty : "&C" + definition.Center) +
        (definition.Right is null ? string.Empty : "&R" + definition.Right);
}
