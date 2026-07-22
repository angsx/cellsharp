namespace CellSharp.Internal;

internal sealed class WorksheetSettingsDefinition
{
    internal WorksheetSettingsDefinition(
        bool autoFilter,
        int? freezeRows,
        int? freezeColumns,
        bool? landscape,
        int? fitToWidth,
        int? fitToHeight,
        bool repeatHeaderRowOnPrint,
        bool printGridlines)
    {
        AutoFilter = autoFilter;
        FreezeRows = freezeRows;
        FreezeColumns = freezeColumns;
        Landscape = landscape;
        FitToWidth = fitToWidth;
        FitToHeight = fitToHeight;
        RepeatHeaderRowOnPrint = repeatHeaderRowOnPrint;
        PrintGridlines = printGridlines;
    }

    internal bool AutoFilter { get; }

    internal int? FreezeRows { get; }

    internal int? FreezeColumns { get; }

    internal bool? Landscape { get; }

    internal int? FitToWidth { get; }

    internal int? FitToHeight { get; }

    internal bool RepeatHeaderRowOnPrint { get; }

    internal bool PrintGridlines { get; }

    internal ResolvedWorksheetSettings Resolve(ExcelWriteOptions options) => new(
        AutoFilter,
        FreezeRows ?? (options.FreezeHeaderRow ? 1 : 0),
        FreezeColumns ?? 0,
        Landscape,
        FitToWidth,
        FitToHeight,
        RepeatHeaderRowOnPrint,
        PrintGridlines);

    internal static ResolvedWorksheetSettings From(ExcelWriteOptions options) => new(
        false,
        options.FreezeHeaderRow ? 1 : 0,
        0,
        null,
        null,
        null,
        false,
        false);
}

internal sealed class ResolvedWorksheetSettings
{
    internal ResolvedWorksheetSettings(
        bool autoFilter,
        int freezeRows,
        int freezeColumns,
        bool? landscape,
        int? fitToWidth,
        int? fitToHeight,
        bool repeatHeaderRowOnPrint,
        bool printGridlines)
    {
        AutoFilter = autoFilter;
        FreezeRows = freezeRows;
        FreezeColumns = freezeColumns;
        Landscape = landscape;
        FitToWidth = fitToWidth;
        FitToHeight = fitToHeight;
        RepeatHeaderRowOnPrint = repeatHeaderRowOnPrint;
        PrintGridlines = printGridlines;
    }

    internal bool AutoFilter { get; }

    internal int FreezeRows { get; }

    internal int FreezeColumns { get; }

    internal bool? Landscape { get; }

    internal int? FitToWidth { get; }

    internal int? FitToHeight { get; }

    internal bool RepeatHeaderRowOnPrint { get; }

    internal bool PrintGridlines { get; }

    internal bool HasFitToPage => FitToWidth is not null || FitToHeight is not null;
}
