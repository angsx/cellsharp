using CellSharp.Internal;

namespace CellSharp;

/// <summary>Horizontal alignment for arbitrary worksheet cells.</summary>
public enum ExcelCellHorizontalAlignment
{
    /// <summary>Uses Excel's general alignment.</summary>
    General,
    /// <summary>Aligns content left.</summary>
    Left,
    /// <summary>Centers content.</summary>
    Center,
    /// <summary>Aligns content right.</summary>
    Right,
    /// <summary>Fills the cell width by repeating the content.</summary>
    Fill,
    /// <summary>Justifies content across the cell.</summary>
    Justify,
    /// <summary>Centers content across adjacent empty cells without merging.</summary>
    CenterContinuous,
    /// <summary>Distributes content across the cell width.</summary>
    Distributed,
}
/// <summary>Vertical alignment for arbitrary worksheet cells.</summary>
public enum ExcelVerticalAlignment
{
    /// <summary>Aligns content to the top.</summary>
    Top,
    /// <summary>Centers content vertically.</summary>
    Center,
    /// <summary>Aligns content to the bottom.</summary>
    Bottom,
    /// <summary>Justifies content vertically.</summary>
    Justify,
    /// <summary>Distributes content vertically.</summary>
    Distributed,
}
/// <summary>Underline decoration for cell text.</summary>
public enum ExcelUnderline
{
    /// <summary>Removes underline decoration.</summary>
    None,
    /// <summary>Uses a single underline.</summary>
    Single,
    /// <summary>Uses a double underline.</summary>
    Double,
}
/// <summary>Line styles supported by worksheet borders.</summary>
public enum ExcelBorderStyle
{
    /// <summary>Removes the border side.</summary>
    None,
    /// <summary>Uses a hairline border.</summary>
    Hair,
    /// <summary>Uses a thin border.</summary>
    Thin,
    /// <summary>Uses a medium border.</summary>
    Medium,
    /// <summary>Uses a thick border.</summary>
    Thick,
    /// <summary>Uses a dashed border.</summary>
    Dashed,
    /// <summary>Uses a dotted border.</summary>
    Dotted,
    /// <summary>Uses a double border.</summary>
    Double,
}

/// <summary>Creates a reusable, immutable cell style.</summary>
public sealed class ExcelStyle
{
    private readonly LayoutStyleDefinition _definition;
    private ExcelStyle(LayoutStyleDefinition definition) => _definition = definition;

    /// <summary>Creates a style from a fluent configuration callback.</summary>
    public static ExcelStyle Create(Action<ExcelStyleBuilder> configure)
    {
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelStyleBuilder();
        configure(builder);
        return new ExcelStyle(builder.Build());
    }

    internal LayoutStyleDefinition Definition => _definition;
}

/// <summary>Builds a composable cell style. Only configured properties override an underlying style.</summary>
public sealed class ExcelStyleBuilder
{
    private LayoutStyleDefinition _style = new();
    /// <summary>Makes text bold.</summary>
    public ExcelStyleBuilder Bold(bool value = true) { _style = _style with { Bold = value }; return this; }
    /// <summary>Makes text italic.</summary>
    public ExcelStyleBuilder Italic(bool value = true) { _style = _style with { Italic = value }; return this; }
    /// <summary>Sets text underline.</summary>
    public ExcelStyleBuilder Underline(ExcelUnderline value = ExcelUnderline.Single) { _style = _style with { Underline = value }; return this; }
    /// <summary>Sets strikethrough.</summary>
    public ExcelStyleBuilder Strikethrough(bool value = true) { _style = _style with { Strikethrough = value }; return this; }
    /// <summary>Sets the font size in points.</summary>
    public ExcelStyleBuilder FontSize(double value) { if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value)); _style = _style with { FontSize = value }; return this; }
    /// <summary>Sets the font family name.</summary>
    public ExcelStyleBuilder FontName(string value) { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A font name is required.", nameof(value)); _style = _style with { FontName = value }; return this; }
    /// <summary>Sets font color using #RRGGBB.</summary>
    public ExcelStyleBuilder FontColor(string value) { _style = _style with { FontColor = HexColor.Normalize(value, nameof(value)) }; return this; }
    /// <summary>Sets a solid fill using #RRGGBB.</summary>
    public ExcelStyleBuilder FillColor(string value) { _style = _style with { FillColor = HexColor.Normalize(value, nameof(value)) }; return this; }
    /// <summary>Sets horizontal alignment.</summary>
    public ExcelStyleBuilder Align(ExcelCellHorizontalAlignment value) { if (!Enum.IsDefined(typeof(ExcelCellHorizontalAlignment), value)) throw new ArgumentOutOfRangeException(nameof(value)); _style = _style with { Horizontal = value }; return this; }
    /// <summary>Aligns content left.</summary>
    public ExcelStyleBuilder AlignLeft() => Align(ExcelCellHorizontalAlignment.Left);
    /// <summary>Centers content horizontally.</summary>
    public ExcelStyleBuilder AlignCenter() => Align(ExcelCellHorizontalAlignment.Center);
    /// <summary>Aligns content right.</summary>
    public ExcelStyleBuilder AlignRight() => Align(ExcelCellHorizontalAlignment.Right);
    /// <summary>Sets vertical alignment.</summary>
    public ExcelStyleBuilder VerticalAlign(ExcelVerticalAlignment value) { if (!Enum.IsDefined(typeof(ExcelVerticalAlignment), value)) throw new ArgumentOutOfRangeException(nameof(value)); _style = _style with { Vertical = value }; return this; }
    /// <summary>Aligns content at the top.</summary>
    public ExcelStyleBuilder VerticalAlignTop() => VerticalAlign(ExcelVerticalAlignment.Top);
    /// <summary>Centers content vertically.</summary>
    public ExcelStyleBuilder VerticalAlignCenter() => VerticalAlign(ExcelVerticalAlignment.Center);
    /// <summary>Aligns content at the bottom.</summary>
    public ExcelStyleBuilder VerticalAlignBottom() => VerticalAlign(ExcelVerticalAlignment.Bottom);
    /// <summary>Rotates text between -90 and 90 degrees.</summary>
    public ExcelStyleBuilder TextRotation(int value) { if (value < -90 || value > 90) throw new ArgumentOutOfRangeException(nameof(value), "Text rotation must be between -90 and 90 degrees."); _style = _style with { Rotation = value, VerticalText = false }; return this; }
    /// <summary>Displays text vertically, one character per line.</summary>
    public ExcelStyleBuilder VerticalText() { _style = _style with { VerticalText = true, Rotation = null }; return this; }
    /// <summary>Enables or disables wrapping.</summary>
    public ExcelStyleBuilder WrapText(bool value = true) { _style = _style with { Wrap = value }; return this; }
    /// <summary>Enables or disables shrink-to-fit.</summary>
    public ExcelStyleBuilder ShrinkToFit(bool value = true) { _style = _style with { Shrink = value }; return this; }
    /// <summary>Sets an Excel number format.</summary>
    public ExcelStyleBuilder NumberFormat(string value) { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("A number format is required.", nameof(value)); _style = _style with { NumberFormat = value }; return this; }
    /// <summary>Configures borders.</summary>
    public ExcelStyleBuilder Border(Action<ExcelBorderBuilder> configure) { if (configure is null) throw new ArgumentNullException(nameof(configure)); var b = new ExcelBorderBuilder(_style.Border); configure(b); _style = _style with { Border = b.Build() }; return this; }
    internal LayoutStyleDefinition Build() => _style;
}

/// <summary>Builds borders for a cell or range.</summary>
public sealed class ExcelBorderBuilder
{
    private BorderDefinition _border;
    internal ExcelBorderBuilder(BorderDefinition? border) => _border = border ?? new BorderDefinition();
    /// <summary>Sets all outer sides.</summary>
    public ExcelBorderBuilder Outline(ExcelBorderStyle style, string? color = null) { var side = new BorderSideDefinition(style, color is null ? null : HexColor.Normalize(color, nameof(color))); _border = _border with { Left = side, Right = side, Top = side, Bottom = side }; return this; }
    /// <summary>Sets all inside sides of a range.</summary>
    public ExcelBorderBuilder Inside(ExcelBorderStyle style, string? color = null) { var side = new BorderSideDefinition(style, color is null ? null : HexColor.Normalize(color, nameof(color))); _border = _border with { InsideHorizontal = side, InsideVertical = side }; return this; }
    /// <summary>Sets the left side.</summary>
    public ExcelBorderBuilder Left(ExcelBorderStyle style, string? color = null) { _border = _border with { Left = Side(style, color) }; return this; }
    /// <summary>Sets the right side.</summary>
    public ExcelBorderBuilder Right(ExcelBorderStyle style, string? color = null) { _border = _border with { Right = Side(style, color) }; return this; }
    /// <summary>Sets the top side.</summary>
    public ExcelBorderBuilder Top(ExcelBorderStyle style, string? color = null) { _border = _border with { Top = Side(style, color) }; return this; }
    /// <summary>Sets the bottom side.</summary>
    public ExcelBorderBuilder Bottom(ExcelBorderStyle style, string? color = null) { _border = _border with { Bottom = Side(style, color) }; return this; }
    private static BorderSideDefinition Side(ExcelBorderStyle style, string? color) => new(style, color is null ? null : HexColor.Normalize(color, nameof(color)));
    internal BorderDefinition Build() => _border;
}

/// <summary>Configures arbitrary cells, ranges, rows, and columns in one worksheet.</summary>
public sealed class ExcelWorksheetBuilder
{
    private readonly WorksheetLayoutDefinition _layout = new();
    internal ExcelWorksheetBuilder(string name) => _layout.Name = ExcelSheetName.Validate(name, nameof(name));
    internal WorksheetLayoutDefinition Definition => _layout;
    /// <summary>Configures this worksheet's printable header.</summary>
    public ExcelHeaderFooter Header => new(_layout.Header);
    /// <summary>Configures this worksheet's printable footer.</summary>
    public ExcelHeaderFooter Footer => new(_layout.Footer);
    /// <summary>Gets one cell using an A1 reference.</summary>
    public ExcelRange Cell(string reference) => new(this, ExcelRangeReference.ParseCell(reference));
    /// <summary>Gets one cell using one-based Excel coordinates.</summary>
    public ExcelRange Cell(int row, int column) => new(this, ExcelRangeReference.Cell(row, column));
    /// <summary>Gets a rectangular range using A1 notation.</summary>
    public ExcelRange Range(string reference)
    {
        if (string.IsNullOrWhiteSpace(reference)) throw new ArgumentException("A range reference is required.", nameof(reference));
        var trimmed = reference.Trim();
        return new(this, trimmed.All(character => char.IsLetter(character) || character == ':')
            ? ExcelRangeReference.ParseColumns(trimmed)
            : ExcelRangeReference.Parse(trimmed));
    }
    /// <summary>Gets a rectangular range using one-based Excel coordinates.</summary>
    public ExcelRange Range(int fromRow, int fromColumn, int toRow, int toColumn) => new(this, new ExcelRangeReference(fromRow, fromColumn, toRow, toColumn));
    /// <summary>Merges a rectangular range.</summary>
    public ExcelRange Merge(string reference) => Range(reference).Merge();
    /// <summary>Gets a row selection.</summary>
    public ExcelRow Row(int row) => new(this, row);
    /// <summary>Gets a column selection.</summary>
    public ExcelColumn Column(string column) => new(this, ExcelRangeReference.ParseColumn(column));
    /// <summary>Gets a contiguous column selection.</summary>
    public ExcelColumns Columns(string columns) { var range = ExcelRangeReference.ParseColumns(columns); return new ExcelColumns(this, range.FromColumn, range.ToColumn); }
    /// <summary>Places the schema header at this cell; columns remain contiguous to its right.</summary>
    public ExcelWorksheetBuilder DataStartAt(string reference) { var cell = ExcelRangeReference.ParseCell(reference); _layout.DataStart = cell; return this; }
    /// <summary>Creates a neutral report title using merge, value, and ordinary cell styling.</summary>
    public ExcelRange Title(string text, string reference) => TextComponent(text, reference, ReportComponentDefaults.Title, "title");
    /// <summary>Creates a neutral report section heading using merge, value, and ordinary cell styling.</summary>
    public ExcelRange Section(string text, string reference) => TextComponent(text, reference, ReportComponentDefaults.Section, "section");
    /// <summary>Creates a rendered report note inside cells; this is distinct from a classic Excel <see cref="ExcelRange.Comment(string, string?)"/>.</summary>
    public ExcelRange Note(string text, string reference) => TextComponent(text, reference, ReportComponentDefaults.Note, "note");
    /// <summary>Creates a KPI whose label occupies the top row and whose value occupies all remaining rows.</summary>
    public ExcelKpi Kpi(string label, object? value, string reference)
    {
        if (string.IsNullOrWhiteSpace(label)) throw new ArgumentException("A KPI label is required.", nameof(label));
        var container = ComponentRange(reference, "KPI");
        if (container.Reference.ToRow - container.Reference.FromRow < 1) throw new ArgumentException("A KPI range must contain at least two rows.", nameof(reference));
        var labelRange = Range(container.Reference.FromRow, container.Reference.FromColumn, container.Reference.FromRow, container.Reference.ToColumn);
        var valueRange = Range(container.Reference.FromRow + 1, container.Reference.FromColumn, container.Reference.ToRow, container.Reference.ToColumn);
        MergeIfMultiple(labelRange); MergeIfMultiple(valueRange);
        Cell(labelRange.Reference.FromRow, labelRange.Reference.FromColumn).Value(label);
        Cell(valueRange.Reference.FromRow, valueRange.Reference.FromColumn).Value(value);
        ReportComponentDefaults.KpiContainer(container); ReportComponentDefaults.KpiLabel(labelRange); ReportComponentDefaults.KpiValue(valueRange);
        return new ExcelKpi(container, labelRange, valueRange);
    }
    internal void Apply(ExcelRangeReference range, LayoutStyleDefinition style) => _layout.Styles.Add(new RangeStyleOperation(range, style));
    internal void SetValue(ExcelRangeReference range, object? value) => _layout.Values.Add(new RangeValueOperation(range, value, null));
    internal void SetFormula(ExcelRangeReference range, string formula) => _layout.Values.Add(new RangeValueOperation(range, null, FormulaNormalizer.Normalize(formula)));
    internal void SetHyperlink(ExcelRangeReference range, string target)
    {
        if (string.IsNullOrWhiteSpace(target)) throw new ArgumentException("A hyperlink target is required.", nameof(target));
        if (target.StartsWith("#", StringComparison.Ordinal))
        {
            if (string.IsNullOrWhiteSpace(target.Substring(1))) throw new ArgumentException("An internal hyperlink must include a worksheet location after '#'.", nameof(target));
        }
        else if (!Uri.TryCreate(target, UriKind.Absolute, out var uri)
            || uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps)
        {
            throw new ArgumentException("An external hyperlink must use an absolute HTTP or HTTPS URL.", nameof(target));
        }

        _layout.Hyperlinks.Add(new HyperlinkOperation(range, target));
    }
    internal void AddConditionalFormat(ConditionalFormattingRuleDefinition definition) => _layout.ConditionalFormats.Add(definition);
    internal void AddDefinedName(ExcelRangeReference range, string name) { ExcelDefinedName.Validate(name, nameof(name)); _layout.DefinedNames.Add(new DefinedNameDefinition(name, range)); }
    internal void AddComment(ExcelRangeReference range, string text, string? author)
    {
        if (!range.IsCell) throw new InvalidOperationException("A comment can only be attached to a single cell.");
        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException("Comment text must not be empty.", nameof(text));
        if (_layout.Merges.Any(merge => merge.Contains(range) && !merge.IsTopLeft(range))) throw new InvalidOperationException("A comment in a merged range must be attached to its top-left cell.");
        var actualAuthor = string.IsNullOrWhiteSpace(author) ? "CellSharp" : author!.Trim();
        if (_layout.Comments.ContainsKey(range)) throw new InvalidOperationException($"Cell {range} already has a comment.");
        _layout.Comments.Add(range, new CommentDefinition(text, actualAuthor));
    }
    internal HeaderFooterDefinition HeaderDefinition => _layout.Header;
    internal HeaderFooterDefinition FooterDefinition => _layout.Footer;
    internal void AddRowPageBreak(int row) => _layout.RowPageBreaks.Add(row);
    internal void AddColumnPageBreak(int column) => _layout.ColumnPageBreaks.Add(column);
    /// <summary>Adds a PNG or JPEG image from a file path. The image data is read immediately.</summary>
    public ExcelImage AddImage(string path) { var image = ExcelImageData.FromPath(path); _layout.Images.Add(image); return new ExcelImage(image); }
    /// <summary>Adds a PNG or JPEG image from a caller-owned readable stream. The image data is read immediately and the stream remains open.</summary>
    public ExcelImage AddImage(Stream stream, ExcelImageFormat format) { if (!Enum.IsDefined(typeof(ExcelImageFormat), format)) throw new ArgumentOutOfRangeException(nameof(format)); var image = ExcelImageData.FromStream(stream, format); _layout.Images.Add(image); return new ExcelImage(image); }
    internal void MergeRange(ExcelRangeReference range) => _layout.AddMerge(range);
    internal void SetRow(int row, double? height, bool? hidden) { if (row < 1 || row > ExcelRangeReference.MaximumRows) throw new ArgumentOutOfRangeException(nameof(row)); var current = _layout.Rows.TryGetValue(row, out var value) ? value : new RowLayoutDefinition(null, null); _layout.Rows[row] = new RowLayoutDefinition(height ?? current.Height, hidden ?? current.Hidden); }
    internal void SetColumns(int from, int to, double? width, bool? hidden, bool? autoFit = null) { for (var column = from; column <= to; column++) { var current = _layout.Columns.TryGetValue(column, out var value) ? value : new ColumnLayoutDefinition(null, null); _layout.Columns[column] = new ColumnLayoutDefinition(width ?? current.Width, hidden ?? current.Hidden, autoFit ?? current.AutoFit); } }
    internal void ApplyRow(int row, LayoutStyleDefinition style) => _layout.RowStyles.Add(new RowStyleOperation(row, style));
    internal void ApplyColumns(int from, int to, LayoutStyleDefinition style) => _layout.ColumnStyles.Add(new ColumnStyleOperation(from, to, style));
    internal void AutoFitColumns(int from, int to) => SetColumns(from, to, null, null, true);

    private ExcelRange TextComponent(string text, string reference, Action<ExcelRange> defaults, string kind)
    {
        if (string.IsNullOrWhiteSpace(text)) throw new ArgumentException($"A {kind} text is required.", nameof(text));
        var range = ComponentRange(reference, kind);
        MergeIfMultiple(range);
        Cell(range.Reference.FromRow, range.Reference.FromColumn).Value(text);
        defaults(range);
        return range;
    }
    private ExcelRange ComponentRange(string reference, string kind)
    {
        var range = Range(reference);
        if (range.Reference.ToRow == ExcelRangeReference.MaximumRows) throw new ArgumentException($"A {kind} range must have a finite row boundary.", nameof(reference));
        return range;
    }
    private static void MergeIfMultiple(ExcelRange range) { if (!range.Reference.IsCell) range.Merge(); }
}

/// <summary>A selected cell or range.</summary>
public sealed class ExcelRange
{
    private readonly ExcelWorksheetBuilder _sheet; private readonly ExcelRangeReference _range;
    internal ExcelRange(ExcelWorksheetBuilder sheet, ExcelRangeReference range) { _sheet = sheet; _range = range; }
    internal ExcelRangeReference Reference => _range;
    /// <summary>Merges this range.</summary>
    public ExcelRange Merge() { _sheet.MergeRange(_range); return this; }
    /// <summary>Sets the value in the top-left cell. For a non-merged range all cells receive the value.</summary>
    public ExcelRange Value(object? value) { _sheet.SetValue(_range, value); return this; }
    /// <summary>Sets a native Excel formula in the top-left cell of a merged range, or in every cell of an unmerged range.</summary>
    public ExcelRange Formula(string formula) { _sheet.SetFormula(_range, formula); return this; }
    /// <summary>Sets an absolute HTTP(S) URL or an internal <c>#Sheet!A1</c> hyperlink.</summary>
    public ExcelRange Hyperlink(string target) { _sheet.SetHyperlink(_range, target); return this; }
    /// <summary>Creates a workbook-scoped name for this absolute worksheet range.</summary>
    public ExcelRange Name(string name) { _sheet.AddDefinedName(_range, name); return this; }
    /// <summary>Adds a classic Excel note/comment to this cell. The default author is CellSharp.</summary>
    public ExcelRange Comment(string text, string? author = null) { _sheet.AddComment(_range, text, author); return this; }
    /// <summary>Starts a native Excel conditional-formatting rule for this range.</summary>
    public ExcelConditionalFormatBuilder ConditionalFormat() => new(_sheet, _range);
    /// <summary>Applies a composable style.</summary>
    public ExcelRange Style(Action<ExcelStyleBuilder> configure) { if (configure is null) throw new ArgumentNullException(nameof(configure)); var builder = new ExcelStyleBuilder(); configure(builder); _sheet.Apply(_range, builder.Build()); return this; }
    /// <summary>Applies a reusable style.</summary>
    public ExcelRange Style(ExcelStyle style) { if (style is null) throw new ArgumentNullException(nameof(style)); _sheet.Apply(_range, style.Definition); return this; }
    /// <summary>Makes this range bold.</summary>
    public ExcelRange Bold() => Style(s => s.Bold());
    /// <summary>Centers this range horizontally.</summary>
    public ExcelRange AlignCenter() => Style(s => s.AlignCenter());
    /// <summary>Sets the Excel number format for this range.</summary>
    public ExcelRange NumberFormat(string value) => Style(s => s.NumberFormat(value));
    /// <summary>Estimates widths for the columns covered by this range from materialized cells only.</summary>
    public ExcelRange AutoFitColumns() { _sheet.AutoFitColumns(_range.FromColumn, _range.ToColumn); return this; }
}

/// <summary>A selected worksheet row.</summary>
public sealed class ExcelRow { private readonly ExcelWorksheetBuilder _sheet; private readonly int _row; internal ExcelRow(ExcelWorksheetBuilder sheet, int row) { if (row < 1 || row > ExcelRangeReference.MaximumRows) throw new ArgumentOutOfRangeException(nameof(row)); _sheet = sheet; _row = row; }
    /// <summary>Sets the row height in points.</summary>
    public ExcelRow Height(double value) { if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value)); _sheet.SetRow(_row, value, null); return this; }
    /// <summary>Hides this row.</summary>
    public ExcelRow Hidden(bool value = true) { _sheet.SetRow(_row, null, value); return this; }
    /// <summary>Applies a composable style to materialized cells in this row.</summary>
    public ExcelRow Style(Action<ExcelStyleBuilder> configure) { if (configure is null) throw new ArgumentNullException(nameof(configure)); var builder = new ExcelStyleBuilder(); configure(builder); _sheet.ApplyRow(_row, builder.Build()); return this; }
    /// <summary>Inserts a native horizontal page break after this worksheet row.</summary>
    public ExcelRow PageBreakAfter() { _sheet.AddRowPageBreak(_row); return this; } }
/// <summary>A selected worksheet column.</summary>
public sealed class ExcelColumn { private readonly ExcelWorksheetBuilder _sheet; private readonly int _column; internal ExcelColumn(ExcelWorksheetBuilder sheet, int column) { _sheet = sheet; _column = column; }
    /// <summary>Sets the column width.</summary>
    public ExcelColumn Width(double value) { if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value)); _sheet.SetColumns(_column, _column, value, null); return this; }
    /// <summary>Hides this column.</summary>
    public ExcelColumn Hidden(bool value = true) { _sheet.SetColumns(_column, _column, null, value); return this; }
    /// <summary>Estimates this column's width from materialized cells only.</summary>
    public ExcelColumn AutoFit() { _sheet.AutoFitColumns(_column, _column); return this; }
    /// <summary>Applies a composable style to materialized cells in this column.</summary>
    public ExcelColumn Style(Action<ExcelStyleBuilder> configure) { if (configure is null) throw new ArgumentNullException(nameof(configure)); var builder = new ExcelStyleBuilder(); configure(builder); _sheet.ApplyColumns(_column, _column, builder.Build()); return this; }
    /// <summary>Inserts a native vertical page break after this worksheet column.</summary>
    public ExcelColumn PageBreakAfter() { _sheet.AddColumnPageBreak(_column); return this; } }
/// <summary>A contiguous set of worksheet columns.</summary>
public sealed class ExcelColumns { private readonly ExcelWorksheetBuilder _sheet; private readonly int _from; private readonly int _to; internal ExcelColumns(ExcelWorksheetBuilder sheet, int from, int to) { _sheet = sheet; _from = from; _to = to; }
    /// <summary>Sets widths for all selected columns.</summary>
    public ExcelColumns Width(double value) { if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) throw new ArgumentOutOfRangeException(nameof(value)); _sheet.SetColumns(_from, _to, value, null); return this; }
    /// <summary>Hides all selected columns.</summary>
    public ExcelColumns Hidden(bool value = true) { _sheet.SetColumns(_from, _to, null, value); return this; }
    /// <summary>Estimates selected column widths from materialized cells only.</summary>
    public ExcelColumns AutoFit() { _sheet.AutoFitColumns(_from, _to); return this; }
    /// <summary>Applies a composable style to materialized cells in these columns.</summary>
    public ExcelColumns Style(Action<ExcelStyleBuilder> configure) { if (configure is null) throw new ArgumentNullException(nameof(configure)); var builder = new ExcelStyleBuilder(); configure(builder); _sheet.ApplyColumns(_from, _to, builder.Build()); return this; } }
