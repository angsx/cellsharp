using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxPageBreakWriter
{
    internal static void Apply(IReadOnlyList<WorksheetValidationContext> worksheets)
    {
        foreach (var context in worksheets)
        {
            var layout = context.Layout;
            if (layout is null) continue;
            if (layout.RowPageBreaks.Count > 0)
            {
                var breaks = new RowBreaks { Count = (uint)layout.RowPageBreaks.Count, ManualBreakCount = (uint)layout.RowPageBreaks.Count };
                foreach (var row in layout.RowPageBreaks.OrderBy(value => value)) breaks.AppendChild(new Break { Id = (uint)row, Min = 0U, Max = (uint)(ExcelRangeReference.MaximumColumns - 1), ManualPageBreak = true });
                context.Worksheet.AppendChild(breaks);
            }
            if (layout.ColumnPageBreaks.Count > 0)
            {
                var breaks = new ColumnBreaks { Count = (uint)layout.ColumnPageBreaks.Count, ManualBreakCount = (uint)layout.ColumnPageBreaks.Count };
                foreach (var column in layout.ColumnPageBreaks.OrderBy(value => value)) breaks.AppendChild(new Break { Id = (uint)column, Min = 0U, Max = (uint)(ExcelRangeReference.MaximumRows - 1), ManualPageBreak = true });
                context.Worksheet.AppendChild(breaks);
            }
        }
    }
}
