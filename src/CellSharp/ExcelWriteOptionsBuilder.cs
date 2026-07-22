using CellSharp.Internal;

namespace CellSharp;

/// <summary>Configures presentation options for one XLSX write operation.</summary>
public sealed class ExcelWriteOptionsBuilder
{
    private ExcelTheme _theme = ExcelTheme.Default;
    private ExcelStyleTemplate? _template;
    private HeaderStyleOverride? _headerStyle;
    private bool _autoFitColumns;
    private bool _freezeHeaderRow;
    private bool _alternatingRows;

    /// <summary>Applies a built-in workbook presentation theme.</summary>
    public ExcelWriteOptionsBuilder Theme(ExcelTheme theme)
    {
        if (!Enum.IsDefined(typeof(ExcelTheme), theme))
        {
            throw new ArgumentOutOfRangeException(nameof(theme));
        }

        _theme = theme;
        return this;
    }

    /// <summary>Uses a reusable custom visual template as the export's base style.</summary>
    public ExcelWriteOptionsBuilder Template(ExcelStyleTemplate template)
    {
        if (template is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(template);
#else
            throw new ArgumentNullException(nameof(template));
#endif
        }

        _template = template;
        return this;
    }

    /// <summary>Overrides selected header style values from the current theme.</summary>
    public ExcelWriteOptionsBuilder HeaderStyle(Action<ExcelHeaderStyleBuilder> configure)
    {
        if (configure is null)
        {
#if NET8_0_OR_GREATER
            ArgumentNullException.ThrowIfNull(configure);
#else
            throw new ArgumentNullException(nameof(configure));
#endif
        }

        var builder = new ExcelHeaderStyleBuilder();
        configure(builder);
        _headerStyle = builder.Build();
        return this;
    }

    /// <summary>Estimates column widths from exported text, subject to built-in bounds.</summary>
    public ExcelWriteOptionsBuilder AutoFitColumns()
    {
        _autoFitColumns = true;
        return this;
    }

    /// <summary>Freezes the first worksheet row.</summary>
    public ExcelWriteOptionsBuilder FreezeHeaderRow()
    {
        _freezeHeaderRow = true;
        return this;
    }

    /// <summary>Applies the active theme or template's alternate background to every second data row.</summary>
    public ExcelWriteOptionsBuilder AlternatingRows()
    {
        _alternatingRows = true;
        return this;
    }

    internal ExcelWriteOptions Build() => new(
        _theme,
        _template,
        _headerStyle,
        _autoFitColumns,
        _freezeHeaderRow,
        _alternatingRows);
}

/// <summary>Configures selected visual properties of the worksheet header.</summary>
public sealed class ExcelHeaderStyleBuilder
{
    private bool? _bold;
    private string? _background;
    private string? _foreground;

    /// <summary>Makes header text bold.</summary>
    public ExcelHeaderStyleBuilder Bold()
    {
        _bold = true;
        return this;
    }

    /// <summary>Sets a header background using a #RRGGBB color.</summary>
    public ExcelHeaderStyleBuilder Background(string color)
    {
        _background = HexColor.Normalize(color, nameof(color));
        return this;
    }

    /// <summary>Sets a header foreground using a #RRGGBB color.</summary>
    public ExcelHeaderStyleBuilder Foreground(string color)
    {
        _foreground = HexColor.Normalize(color, nameof(color));
        return this;
    }

    internal HeaderStyleOverride Build() => new(_bold, _background, _foreground);
}
