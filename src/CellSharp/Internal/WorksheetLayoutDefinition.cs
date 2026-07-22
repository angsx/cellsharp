namespace CellSharp.Internal;

internal sealed class WorksheetLayoutDefinition
{
    internal string Name { get; set; } = ExcelSheetName.Default;
    internal ExcelRangeReference? DataStart { get; set; }
    internal List<RangeStyleOperation> Styles { get; } = new();
    internal List<RowStyleOperation> RowStyles { get; } = new();
    internal List<ColumnStyleOperation> ColumnStyles { get; } = new();
    internal List<RangeValueOperation> Values { get; } = new();
    internal List<HyperlinkOperation> Hyperlinks { get; } = new();
    internal List<ConditionalFormattingRuleDefinition> ConditionalFormats { get; } = new();
    internal List<DefinedNameDefinition> DefinedNames { get; } = new();
    internal Dictionary<ExcelRangeReference, CommentDefinition> Comments { get; } = new();
    internal List<WorksheetImageDefinition> Images { get; } = new();
    internal List<ExcelRangeReference> Merges { get; } = new();
    internal HeaderFooterDefinition Header { get; } = new();
    internal HeaderFooterDefinition Footer { get; } = new();
    internal HashSet<int> RowPageBreaks { get; } = new();
    internal HashSet<int> ColumnPageBreaks { get; } = new();
    internal Dictionary<int, RowLayoutDefinition> Rows { get; } = new();
    internal Dictionary<int, ColumnLayoutDefinition> Columns { get; } = new();
    internal void AddMerge(ExcelRangeReference range)
    {
        foreach (var current in Merges)
        {
            if (current.Equals(range)) return;
            if (current.Overlaps(range)) throw new InvalidOperationException($"Range {range} overlaps the merged range {current}.");
        }
        if (Comments.Keys.Any(comment => range.Contains(comment) && !range.IsTopLeft(comment)))
            throw new InvalidOperationException($"A comment in merged range {range} must be attached to its top-left cell.");
        Merges.Add(range);
    }
}

internal sealed record LayoutStyleDefinition(
    bool? Bold = null, bool? Italic = null, ExcelUnderline? Underline = null, bool? Strikethrough = null,
    double? FontSize = null, string? FontName = null, string? FontColor = null, string? FillColor = null,
    ExcelCellHorizontalAlignment? Horizontal = null, ExcelVerticalAlignment? Vertical = null, int? Rotation = null,
    bool? VerticalText = null, bool? Wrap = null, bool? Shrink = null, string? NumberFormat = null,
    BorderDefinition? Border = null)
{
    internal LayoutStyleDefinition Overlay(LayoutStyleDefinition value) => new(
        value.Bold ?? Bold, value.Italic ?? Italic, value.Underline ?? Underline, value.Strikethrough ?? Strikethrough,
        value.FontSize ?? FontSize, value.FontName ?? FontName, value.FontColor ?? FontColor, value.FillColor ?? FillColor,
        value.Horizontal ?? Horizontal, value.Vertical ?? Vertical, value.Rotation ?? Rotation,
        value.VerticalText ?? VerticalText, value.Wrap ?? Wrap, value.Shrink ?? Shrink, value.NumberFormat ?? NumberFormat,
        Border?.Overlay(value.Border) ?? value.Border);
}
internal sealed record BorderDefinition(BorderSideDefinition? Left = null, BorderSideDefinition? Right = null, BorderSideDefinition? Top = null, BorderSideDefinition? Bottom = null, BorderSideDefinition? InsideHorizontal = null, BorderSideDefinition? InsideVertical = null)
{ internal BorderDefinition Overlay(BorderDefinition? value) => value is null ? this : new(value.Left ?? Left, value.Right ?? Right, value.Top ?? Top, value.Bottom ?? Bottom, value.InsideHorizontal ?? InsideHorizontal, value.InsideVertical ?? InsideVertical); }
internal sealed record BorderSideDefinition(ExcelBorderStyle Style, string? Color);
internal sealed record RangeStyleOperation(ExcelRangeReference Range, LayoutStyleDefinition Style);
internal sealed record RangeValueOperation(ExcelRangeReference Range, object? Value, string? Formula);
internal sealed record HyperlinkOperation(ExcelRangeReference Range, string Target);
internal sealed record DefinedNameDefinition(string Name, ExcelRangeReference Range);
internal sealed record CommentDefinition(string Text, string Author);
internal sealed class HeaderFooterDefinition
{
    internal string? Left { get; private set; }
    internal string? Center { get; private set; }
    internal string? Right { get; private set; }
    internal void SetLeft(string value) => Left = value;
    internal void SetCenter(string value) => Center = value;
    internal void SetRight(string value) => Right = value;
    internal bool HasValue => Left is not null || Center is not null || Right is not null;
}
internal enum ConditionalFormattingRuleKind
{
    GreaterThan, GreaterThanOrEqual, LessThan, LessThanOrEqual, EqualTo, NotEqualTo,
    Between, NotBetween, Formula, ContainsText, BeginsWith, EndsWith, DuplicateValues,
    UniqueValues, Blanks, NonBlanks,
}
internal sealed class ConditionalFormattingRuleDefinition
{
    internal ConditionalFormattingRuleDefinition(ExcelRangeReference range, ConditionalFormattingRuleKind kind, IReadOnlyList<string>? formulas = null, string? text = null)
    { Range = range; Kind = kind; Formulas = formulas ?? Array.Empty<string>(); Text = text; }
    internal ExcelRangeReference Range { get; }
    internal ConditionalFormattingRuleKind Kind { get; }
    internal IReadOnlyList<string> Formulas { get; }
    internal string? Text { get; }
    internal LayoutStyleDefinition? Style { get; private set; }
    internal bool StopIfTrue { get; private set; }
    internal void SetStyle(LayoutStyleDefinition style)
    { if (Style is not null) throw new InvalidOperationException("A conditional formatting style can only be configured once."); Style = style; }
    internal void SetStopIfTrue()
    { if (StopIfTrue) throw new InvalidOperationException("StopIfTrue is already configured for this conditional formatting rule."); StopIfTrue = true; }
}
internal sealed record RowStyleOperation(int Row, LayoutStyleDefinition Style);
internal sealed record ColumnStyleOperation(int FromColumn, int ToColumn, LayoutStyleDefinition Style);
internal sealed record RowLayoutDefinition(double? Height, bool? Hidden);
internal sealed record ColumnLayoutDefinition(double? Width, bool? Hidden, bool? AutoFit = null);

internal readonly struct ExcelRangeReference : IEquatable<ExcelRangeReference>
{
    internal const int MaximumRows = 1048576;
    internal const int MaximumColumns = 16384;
    internal ExcelRangeReference(int fromRow, int fromColumn, int toRow, int toColumn)
    {
        Validate(fromRow, fromColumn); Validate(toRow, toColumn);
        if (fromRow > toRow || fromColumn > toColumn) throw new ArgumentException("A range's start must not be after its end.");
        FromRow = fromRow; FromColumn = fromColumn; ToRow = toRow; ToColumn = toColumn;
    }
    internal int FromRow { get; } internal int FromColumn { get; } internal int ToRow { get; } internal int ToColumn { get; }
    internal bool IsCell => FromRow == ToRow && FromColumn == ToColumn;
    internal bool Contains(ExcelRangeReference value) => FromRow <= value.FromRow && ToRow >= value.ToRow && FromColumn <= value.FromColumn && ToColumn >= value.ToColumn;
    internal bool IsTopLeft(ExcelRangeReference value) => value.IsCell && value.FromRow == FromRow && value.FromColumn == FromColumn;
    internal static ExcelRangeReference Cell(int row, int column) => new(row, column, row, column);
    internal static ExcelRangeReference ParseCell(string value) { var result = Parse(value); if (!result.IsCell) throw new ArgumentException("A single cell reference is required.", nameof(value)); return result; }
    internal static ExcelRangeReference Parse(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Contains('!')) throw new ArgumentException("A local A1 cell or range reference is required.", nameof(value));
        var parts = value.Trim().Split(':'); if (parts.Length is < 1 or > 2) throw new ArgumentException("The range reference is invalid.", nameof(value));
        var first = ParseAddress(parts[0], nameof(value)); var second = parts.Length == 1 ? first : ParseAddress(parts[1], nameof(value));
        return new ExcelRangeReference(first.row, first.column, second.row, second.column);
    }
    internal static int ParseColumn(string value)
    {
        if (string.IsNullOrWhiteSpace(value) || value.Any(c => !char.IsLetter(c))) throw new ArgumentException("A column reference such as A or XFD is required.", nameof(value));
        var column = 0; foreach (var character in value.ToUpperInvariant()) { column = (column * 26) + character - 'A' + 1; if (column > MaximumColumns) throw new ArgumentOutOfRangeException(nameof(value), $"Column reference {value} exceeds the XLSX maximum column XFD."); } return column;
    }
    internal static ExcelRangeReference ParseColumns(string value) { var parts = (value ?? string.Empty).Trim().Split(':'); if (parts.Length is < 1 or > 2) throw new ArgumentException("A column range is invalid.", nameof(value)); var start = ParseColumn(parts[0]); var end = parts.Length == 1 ? start : ParseColumn(parts[1]); return new ExcelRangeReference(1, start, MaximumRows, end); }
    internal string CellReference(int row, int column) => $"{ColumnName(column)}{row}";
    internal static string ColumnName(int value) { var name = string.Empty; while (value > 0) { value--; name = (char)('A' + (value % 26)) + name; value /= 26; } return name; }
    public override string ToString() => IsCell ? CellReference(FromRow, FromColumn) : $"{CellReference(FromRow, FromColumn)}:{CellReference(ToRow, ToColumn)}";
    internal bool Overlaps(ExcelRangeReference other) => FromRow <= other.ToRow && ToRow >= other.FromRow && FromColumn <= other.ToColumn && ToColumn >= other.FromColumn;
    public bool Equals(ExcelRangeReference other) => FromRow == other.FromRow && FromColumn == other.FromColumn && ToRow == other.ToRow && ToColumn == other.ToColumn;
    public override bool Equals(object? obj) => obj is ExcelRangeReference other && Equals(other);
    public override int GetHashCode() { unchecked { var hash = FromRow; hash = (hash * 397) ^ FromColumn; hash = (hash * 397) ^ ToRow; return (hash * 397) ^ ToColumn; } }
    private static (int row, int column) ParseAddress(string value, string parameterName) { var source = value.Trim().Replace("$", string.Empty); var split = 0; while (split < source.Length && char.IsLetter(source[split])) split++; if (split == 0 || split == source.Length || !int.TryParse(source.Substring(split), out var row)) throw new ArgumentException("The A1 reference is invalid.", parameterName); return (row, ParseColumn(source.Substring(0, split))); }
    private static void Validate(int row, int column) { if (row < 1 || row > MaximumRows) throw new ArgumentOutOfRangeException(nameof(row), $"Row must be between 1 and {MaximumRows}."); if (column < 1 || column > MaximumColumns) throw new ArgumentOutOfRangeException(nameof(column), $"Column must be between 1 and {MaximumColumns}."); }
}

internal static class FormulaNormalizer
{
    internal static string Normalize(string formula)
    { if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("A formula must not be empty.", nameof(formula)); var value = formula.Trim(); if (value.StartsWith("==", StringComparison.Ordinal)) throw new ArgumentException("A formula can have at most one leading '='.", nameof(formula)); return value[0] == '=' ? value.Substring(1) : value; }
}
