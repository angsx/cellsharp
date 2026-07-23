using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal sealed class WorkbookStyleCatalog
{
    private const string DefaultDateFormat = "yyyy-mm-dd hh:mm:ss";
    private const string DefaultDecimalFormat = "0.00";
    private readonly Fonts _fonts = new(new Font()) { Count = 1U };
    private readonly Fills _fills = new(
        new Fill(new PatternFill { PatternType = PatternValues.None }),
        new Fill(new PatternFill { PatternType = PatternValues.Gray125 })) { Count = 2U };
    private readonly Borders _borders = new(new Border()) { Count = 1U };
    private readonly CellFormats _formats = new(new CellFormat()) { Count = 1U };
    private readonly DifferentialFormats _differentialFormats = new();
    private readonly NumberingFormats _numberingFormats = new();
    private readonly Dictionary<CellStyleDefinition, uint> _styles = new();
    private readonly Dictionary<string, uint> _fontIds = new();
    private readonly Dictionary<string, uint> _fillIds = new();
    private readonly Dictionary<string, uint> _borderIds = new();
    private readonly Dictionary<string, uint> _styleIds = new();
    private readonly Dictionary<string, uint> _numberFormatIds = new(StringComparer.Ordinal);
    private readonly Dictionary<uint, LayoutStyleDefinition> _resolvedStyles = new();
    private readonly Dictionary<string, uint> _differentialStyleIds = new();
    private readonly WorkbookStylesPart _stylesPart;
    internal WorkbookStyleCatalog(WorkbookPart workbookPart)
    {
        _stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
        _stylesPart.Stylesheet = new Stylesheet(
            _numberingFormats,
            _fonts,
            _fills,
            _borders,
            new CellStyleFormats(new CellFormat()) { Count = 1U },
            _formats,
            new CellStyles(new CellStyle { Name = "Normal", FormatId = 0U, BuiltinId = 0U }) { Count = 1U },
            _differentialFormats);
    }

    internal uint HeaderStyleIndex(ExcelWriteOptions options)
    {
        var template = options.Template ?? WorkbookTheme.For(options.Theme);
        var header = new CellStyleDefinition(
            template.HeaderFontName,
            template.HeaderFontSize,
            template.HeaderBold,
            HexColor.Normalize(template.HeaderTextColor, nameof(template.HeaderTextColor)),
            HexColor.Normalize(template.HeaderBackgroundColor, nameof(template.HeaderBackgroundColor)),
            HexColor.Normalize(template.BorderColor, nameof(template.BorderColor)),
            ExcelHorizontalAlignment.Center);
        return Register(header.With(options.HeaderStyle));
    }

    internal uint DataStyleIndex(ExportProperty property, object? value, int dataRowIndex, ExcelWriteOptions options)
    {
        return DataStyleIndex(property, value is DateTime, value is decimal, dataRowIndex, options);
    }

    internal uint TemplateDataStyleIndex(ExportProperty property, ExcelWriteOptions options)
    {
        var propertyType = Nullable.GetUnderlyingType(property.CellType) ?? property.CellType;
        return DataStyleIndex(property, propertyType == typeof(DateTime), propertyType == typeof(decimal), 0, options);
    }

    private uint DataStyleIndex(ExportProperty property, bool isDate, bool isDecimal, int dataRowIndex, ExcelWriteOptions options)
    {
        var template = options.Template ?? WorkbookTheme.For(options.Theme);
        var dataStyle = new CellStyleDefinition(
            template.FontName,
            template.FontSize,
            false,
            HexColor.Normalize(template.DataTextColor, nameof(template.DataTextColor)),
            HexColor.Normalize(template.DataBackgroundColor, nameof(template.DataBackgroundColor)),
            HexColor.Normalize(template.BorderColor, nameof(template.BorderColor)),
            ExcelHorizontalAlignment.Left);
        var alternateBackground = HexColor.Normalize(
            template.AlternateRowBackgroundColor ?? template.DataBackgroundColor,
            nameof(template.AlternateRowBackgroundColor));
        var style = options.AlternatingRows && dataRowIndex % 2 == 1
            ? dataStyle.WithBackground(alternateBackground)
            : dataStyle;
        style = style.WithAlignment(property.Alignment);
        var numberFormat = property.Format ?? (isDate ? DefaultDateFormat : isDecimal ? DefaultDecimalFormat : null);
        return Register(style, numberFormat);
    }

    internal void Save()
    {
        _numberingFormats.Count = (uint)_numberingFormats.ChildElements.Count;
        _differentialFormats.Count = (uint)_differentialFormats.ChildElements.Count;
        _stylesPart.Stylesheet!.Save();
    }

    internal uint DifferentialStyleIndex(LayoutStyleDefinition style)
    {
        ValidateDifferentialStyle(style);
        var key = style.ToString();
        if (_differentialStyleIds.TryGetValue(key, out var existing)) return existing;
        var differential = new DifferentialFormat();
        if (style.Bold is not null || style.Italic is not null || style.Underline is not null || style.Strikethrough is not null || style.FontColor is not null)
        {
            var font = new Font();
            if (style.Bold is not null) font.Append(new Bold { Val = style.Bold.Value });
            if (style.Italic is not null) font.Append(new Italic { Val = style.Italic.Value });
            if (style.Strikethrough is not null) font.Append(new Strike { Val = style.Strikethrough.Value });
            if (style.Underline is not null) font.Append(style.Underline.Value switch
            {
                ExcelUnderline.None => new Underline { Val = UnderlineValues.None },
                ExcelUnderline.Single => new Underline(),
                ExcelUnderline.Double => new Underline { Val = UnderlineValues.Double },
                _ => throw new ArgumentOutOfRangeException(nameof(style)),
            });
            if (style.FontColor is not null) font.Append(new Color { Rgb = OpaqueColor(style.FontColor) });
            differential.Append(font);
        }
        if (style.NumberFormat is not null)
            differential.Append(new NumberingFormat { NumberFormatId = NumberFormatId(style.NumberFormat), FormatCode = style.NumberFormat });
        if (style.FillColor is not null)
            differential.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = OpaqueColor(style.FillColor) }) { PatternType = PatternValues.Solid }));
        if (style.Border is not null)
            differential.Append(new Border(DifferentialLeft(style.Border.Left), DifferentialRight(style.Border.Right), DifferentialTop(style.Border.Top), DifferentialBottom(style.Border.Bottom)));
        _differentialFormats.Append(differential);
        var index = (uint)(_differentialFormats.ChildElements.Count - 1);
        _differentialFormats.Count = index + 1U;
        _differentialStyleIds[key] = index;
        return index;
    }

    internal static void ValidateDifferentialStyle(LayoutStyleDefinition style)
    {
        if (style.FontSize is not null || style.FontName is not null || style.Horizontal is not null || style.Vertical is not null || style.Rotation is not null || style.VerticalText is not null || style.Wrap is not null || style.Shrink is not null)
            throw new InvalidOperationException("Conditional formatting styles support font decorations and color, fill, border, and number format only.");
        if (style.Border?.InsideHorizontal is not null || style.Border?.InsideVertical is not null)
            throw new InvalidOperationException("Conditional formatting styles do not support inside borders.");
    }

    private static LeftBorder DifferentialLeft(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };
    private static RightBorder DifferentialRight(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };
    private static TopBorder DifferentialTop(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };
    private static BottomBorder DifferentialBottom(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };

    internal uint LayoutStyleIndex(LayoutStyleDefinition style) => ComposeStyleIndex(null, style);

    /// <summary>Resolves a sparse layout style over an existing cell format, property by property.</summary>
    internal uint ComposeStyleIndex(uint? baseStyleIndex, LayoutStyleDefinition overlay)
    {
        var style = (baseStyleIndex is not null && _resolvedStyles.TryGetValue(baseStyleIndex.Value, out var current)
            ? current
            : DefaultLayoutStyle).Overlay(overlay);
        var fontName = style.FontName!;
        var fontSize = style.FontSize!.Value;
        var fontColor = style.FontColor!;
        var fillColor = style.FillColor!;
        var border = style.Border!;
        var numberFormatId = style.NumberFormat is null ? (uint?)null : NumberFormatId(style.NumberFormat);
        var key = string.Join("|", fontName, fontSize, style.Bold, style.Italic, style.Underline, style.Strikethrough, fontColor, fillColor, border, style.Horizontal, style.Vertical, style.Rotation, style.VerticalText, style.Wrap, style.Shrink, numberFormatId);
        if (_styleIds.TryGetValue(key, out var existing)) return existing;
        var fontKey = $"layout|{fontName}|{fontSize}|{style.Bold}|{style.Italic}|{style.Underline}|{style.Strikethrough}|{fontColor}";
        uint fontId;
        if (!_fontIds.TryGetValue(fontKey, out fontId))
        {
            var font = new Font();
            if (style.Bold == true) font.Append(new Bold());
            if (style.Italic == true) font.Append(new Italic());
            if (style.Strikethrough == true) font.Append(new Strike());
            if (style.Underline == ExcelUnderline.Single) font.Append(new Underline());
            if (style.Underline == ExcelUnderline.Double) font.Append(new Underline { Val = UnderlineValues.Double });
            font.Append(new FontSize { Val = fontSize }, new Color { Rgb = OpaqueColor(fontColor) }, new FontName { Val = fontName });
            _fonts.Append(font); fontId = _fonts.Count!.Value; _fonts.Count = fontId + 1U; _fontIds[fontKey] = fontId;
        }
        var fillId = FillId(fillColor);
        var borderId = LayoutBorderId(border);
        var alignment = new Alignment
        {
            Horizontal = LayoutHorizontal(style.Horizontal), Vertical = LayoutVertical(style.Vertical),
            WrapText = style.Wrap, ShrinkToFit = style.Shrink,
            TextRotation = style.VerticalText == true ? 255U : style.Rotation is null ? null : (uint?)(style.Rotation.Value < 0 ? 90 - style.Rotation.Value : style.Rotation.Value),
        };
        var format = new CellFormat { FontId = fontId, FillId = fillId, BorderId = borderId, ApplyFont = true, ApplyFill = true, ApplyBorder = true, ApplyAlignment = true, Alignment = alignment, NumberFormatId = numberFormatId, ApplyNumberFormat = numberFormatId is not null };
        _formats.Append(format); var index = _formats.Count!.Value; _formats.Count = index + 1U; _styleIds[key] = index; _resolvedStyles[index] = style; return index;
    }

    private static readonly LayoutStyleDefinition DefaultLayoutStyle = new(
        Bold: false, Italic: false, Underline: ExcelUnderline.None, Strikethrough: false,
        FontSize: 11D, FontName: "Calibri", FontColor: "000000", FillColor: "FFFFFF",
        Horizontal: ExcelCellHorizontalAlignment.General, Vertical: ExcelVerticalAlignment.Bottom,
        Rotation: null, VerticalText: false, Wrap: false, Shrink: false, NumberFormat: null,
        Border: new BorderDefinition());

    private uint LayoutBorderId(BorderDefinition border)
    {
        var key = "layout|" + border;
        if (_borderIds.TryGetValue(key, out var existing)) return existing;
        var value = new Border(
            LayoutLeft(border.Left), LayoutRight(border.Right), LayoutTop(border.Top), LayoutBottom(border.Bottom));
        _borders.Append(value); var index = _borders.Count!.Value; _borders.Count = index + 1U; _borderIds[key] = index; return index;
    }

    private static LeftBorder LayoutLeft(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };
    private static RightBorder LayoutRight(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };
    private static TopBorder LayoutTop(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };
    private static BottomBorder LayoutBottom(BorderSideDefinition? side) => new() { Style = side is null ? null : LayoutBorderStyle(side.Style), Color = LayoutBorderColor(side) };
    private static Color? LayoutBorderColor(BorderSideDefinition? side) => side?.Color is null ? null : new Color { Rgb = OpaqueColor(side.Color) };

    private static HorizontalAlignmentValues? LayoutHorizontal(ExcelCellHorizontalAlignment? value) => value switch { null => null, ExcelCellHorizontalAlignment.General => HorizontalAlignmentValues.General, ExcelCellHorizontalAlignment.Left => HorizontalAlignmentValues.Left, ExcelCellHorizontalAlignment.Center => HorizontalAlignmentValues.Center, ExcelCellHorizontalAlignment.Right => HorizontalAlignmentValues.Right, ExcelCellHorizontalAlignment.Fill => HorizontalAlignmentValues.Fill, ExcelCellHorizontalAlignment.Justify => HorizontalAlignmentValues.Justify, ExcelCellHorizontalAlignment.CenterContinuous => HorizontalAlignmentValues.CenterContinuous, ExcelCellHorizontalAlignment.Distributed => HorizontalAlignmentValues.Distributed, _ => throw new ArgumentOutOfRangeException(nameof(value)) };
    private static VerticalAlignmentValues? LayoutVertical(ExcelVerticalAlignment? value) => value switch { null => null, ExcelVerticalAlignment.Top => VerticalAlignmentValues.Top, ExcelVerticalAlignment.Center => VerticalAlignmentValues.Center, ExcelVerticalAlignment.Bottom => VerticalAlignmentValues.Bottom, ExcelVerticalAlignment.Justify => VerticalAlignmentValues.Justify, ExcelVerticalAlignment.Distributed => VerticalAlignmentValues.Distributed, _ => throw new ArgumentOutOfRangeException(nameof(value)) };
    private static BorderStyleValues? LayoutBorderStyle(ExcelBorderStyle value) => value switch { ExcelBorderStyle.None => null, ExcelBorderStyle.Hair => BorderStyleValues.Hair, ExcelBorderStyle.Thin => BorderStyleValues.Thin, ExcelBorderStyle.Medium => BorderStyleValues.Medium, ExcelBorderStyle.Thick => BorderStyleValues.Thick, ExcelBorderStyle.Dashed => BorderStyleValues.Dashed, ExcelBorderStyle.Dotted => BorderStyleValues.Dotted, ExcelBorderStyle.Double => BorderStyleValues.Double, _ => throw new ArgumentOutOfRangeException(nameof(value)) };

    private uint Register(CellStyleDefinition style, string? numberFormat = null)
    {
        if (numberFormat is null && _styles.TryGetValue(style, out var existing)) return existing;
        var borderColor = style.Border;
        var resolved = new LayoutStyleDefinition(
            Bold: style.Bold, Italic: false, Underline: ExcelUnderline.None, Strikethrough: false,
            FontSize: style.FontSize, FontName: style.FontName, FontColor: style.Foreground, FillColor: style.Background,
            Horizontal: style.Alignment switch { ExcelHorizontalAlignment.General => ExcelCellHorizontalAlignment.General, ExcelHorizontalAlignment.Left => ExcelCellHorizontalAlignment.Left, ExcelHorizontalAlignment.Center => ExcelCellHorizontalAlignment.Center, ExcelHorizontalAlignment.Right => ExcelCellHorizontalAlignment.Right, _ => throw new ArgumentOutOfRangeException(nameof(style)) },
            Vertical: ExcelVerticalAlignment.Center, Rotation: null, VerticalText: false, Wrap: false, Shrink: false,
            NumberFormat: numberFormat,
            Border: new BorderDefinition(new BorderSideDefinition(ExcelBorderStyle.Thin, borderColor), new BorderSideDefinition(ExcelBorderStyle.Thin, borderColor), new BorderSideDefinition(ExcelBorderStyle.Thin, borderColor), new BorderSideDefinition(ExcelBorderStyle.Thin, borderColor)));
        var index = ComposeStyleIndex(null, resolved);
        if (numberFormat is null) _styles[style] = index;
        return index;
    }

    private uint NumberFormatId(string format)
    {
        if (_numberFormatIds.TryGetValue(format, out var id))
        {
            return id;
        }

        id = (uint)(164 + _numberFormatIds.Count);
        _numberingFormats.Append(new NumberingFormat { NumberFormatId = id, FormatCode = format });
        _numberFormatIds[format] = id;
        return id;
    }

    private uint FontId(CellStyleDefinition style)
    {
        var key = $"{style.FontName}|{style.FontSize}|{style.Bold}|{style.Foreground}";
        if (_fontIds.TryGetValue(key, out var index))
        {
            return index;
        }

        var font = new Font();
        if (style.Bold)
        {
            font.Append(new Bold());
        }

        font.Append(
            new FontSize { Val = style.FontSize },
            new Color { Rgb = OpaqueColor(style.Foreground) },
            new FontName { Val = style.FontName });

        _fonts.Append(font);
        index = _fonts.Count!.Value;
        _fonts.Count = index + 1U;
        _fontIds[key] = index;
        return index;
    }

    private uint FillId(string color)
    {
        if (_fillIds.TryGetValue(color, out var index))
        {
            return index;
        }

        _fills.Append(new Fill(new PatternFill(new ForegroundColor { Rgb = OpaqueColor(color) }) { PatternType = PatternValues.Solid }));
        index = _fills.Count!.Value;
        _fills.Count = index + 1U;
        _fillIds[color] = index;
        return index;
    }

    private uint BorderId(string color)
    {
        if (_borderIds.TryGetValue(color, out var index))
        {
            return index;
        }

        _borders.Append(new Border(
            new LeftBorder(new Color { Rgb = OpaqueColor(color) }) { Style = BorderStyleValues.Thin },
            new RightBorder(new Color { Rgb = OpaqueColor(color) }) { Style = BorderStyleValues.Thin },
            new TopBorder(new Color { Rgb = OpaqueColor(color) }) { Style = BorderStyleValues.Thin },
            new BottomBorder(new Color { Rgb = OpaqueColor(color) }) { Style = BorderStyleValues.Thin }));
        index = _borders.Count!.Value;
        _borders.Count = index + 1U;
        _borderIds[color] = index;
        return index;
    }

    private static HorizontalAlignmentValues HorizontalAlignment(ExcelHorizontalAlignment alignment) => alignment switch
    {
        ExcelHorizontalAlignment.General => HorizontalAlignmentValues.General,
        ExcelHorizontalAlignment.Left => HorizontalAlignmentValues.Left,
        ExcelHorizontalAlignment.Center => HorizontalAlignmentValues.Center,
        ExcelHorizontalAlignment.Right => HorizontalAlignmentValues.Right,
        _ => throw new ArgumentOutOfRangeException(nameof(alignment)),
    };

    private static string StyleKey(CellStyleDefinition style, uint? numberFormatId) =>
        $"{style.FontName}|{style.FontSize}|{style.Bold}|{style.Foreground}|{style.Background}|{style.Border}|{style.Alignment}|{numberFormatId}";

    private static string OpaqueColor(string rgb) => $"FF{rgb}";
}
