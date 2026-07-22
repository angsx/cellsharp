namespace CellSharp.Internal;

internal sealed class ExcelWriteOptions
{
    internal ExcelWriteOptions(
        ExcelTheme theme,
        ExcelStyleTemplate? template,
        HeaderStyleOverride? headerStyle,
        bool autoFitColumns,
        bool freezeHeaderRow,
        bool alternatingRows)
    {
        Theme = theme;
        Template = template;
        HeaderStyle = headerStyle;
        AutoFitColumns = autoFitColumns;
        FreezeHeaderRow = freezeHeaderRow;
        AlternatingRows = alternatingRows;
    }

    internal ExcelTheme Theme { get; }

    internal ExcelStyleTemplate? Template { get; }

    internal HeaderStyleOverride? HeaderStyle { get; }

    internal bool AutoFitColumns { get; }

    internal bool FreezeHeaderRow { get; }

    internal bool AlternatingRows { get; }

    internal static ExcelWriteOptions Default { get; } = new(ExcelTheme.Default, null, null, false, false, false);
}
