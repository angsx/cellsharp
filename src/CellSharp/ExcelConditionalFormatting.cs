using System.Globalization;
using CellSharp.Internal;

namespace CellSharp;

/// <summary>Selects one native Excel conditional-formatting rule for a range.</summary>
public sealed class ExcelConditionalFormatBuilder
{
    private readonly ExcelWorksheetBuilder _sheet;
    private readonly ExcelRangeReference _range;

    internal ExcelConditionalFormatBuilder(ExcelWorksheetBuilder sheet, ExcelRangeReference range) { _sheet = sheet; _range = range; }

    /// <summary>Formats cells whose value is greater than the supplied value.</summary>
    public ExcelConditionalFormatRuleBuilder GreaterThan(object value) => Add(ConditionalFormattingRuleKind.GreaterThan, Value(value));
    /// <summary>Formats cells whose value is greater than or equal to the supplied value.</summary>
    public ExcelConditionalFormatRuleBuilder GreaterThanOrEqual(object value) => Add(ConditionalFormattingRuleKind.GreaterThanOrEqual, Value(value));
    /// <summary>Formats cells whose value is less than the supplied value.</summary>
    public ExcelConditionalFormatRuleBuilder LessThan(object value) => Add(ConditionalFormattingRuleKind.LessThan, Value(value));
    /// <summary>Formats cells whose value is less than or equal to the supplied value.</summary>
    public ExcelConditionalFormatRuleBuilder LessThanOrEqual(object value) => Add(ConditionalFormattingRuleKind.LessThanOrEqual, Value(value));
    /// <summary>Formats cells whose value equals the supplied value.</summary>
    public ExcelConditionalFormatRuleBuilder EqualTo(object value) => Add(ConditionalFormattingRuleKind.EqualTo, Value(value));
    /// <summary>Formats cells whose value differs from the supplied value.</summary>
    public ExcelConditionalFormatRuleBuilder NotEqualTo(object value) => Add(ConditionalFormattingRuleKind.NotEqualTo, Value(value));
    /// <summary>Formats cells whose value lies between the supplied inclusive bounds.</summary>
    public ExcelConditionalFormatRuleBuilder Between(object minimum, object maximum) => Add(ConditionalFormattingRuleKind.Between, Value(minimum), Value(maximum));
    /// <summary>Formats cells whose value lies outside the supplied inclusive bounds.</summary>
    public ExcelConditionalFormatRuleBuilder NotBetween(object minimum, object maximum) => Add(ConditionalFormattingRuleKind.NotBetween, Value(minimum), Value(maximum));
    /// <summary>Formats cells for which the native Excel formula evaluates to true.</summary>
    public ExcelConditionalFormatRuleBuilder Formula(string formula)
    {
        if (string.IsNullOrWhiteSpace(formula)) throw new ArgumentException("A conditional formatting formula must not be empty.", nameof(formula));
        return Add(ConditionalFormattingRuleKind.Formula, FormulaNormalizer.Normalize(formula));
    }
    /// <summary>Formats cells containing the supplied text.</summary>
    public ExcelConditionalFormatRuleBuilder ContainsText(string value) => Text(ConditionalFormattingRuleKind.ContainsText, value);
    /// <summary>Formats cells beginning with the supplied text.</summary>
    public ExcelConditionalFormatRuleBuilder BeginsWith(string value) => Text(ConditionalFormattingRuleKind.BeginsWith, value);
    /// <summary>Formats cells ending with the supplied text.</summary>
    public ExcelConditionalFormatRuleBuilder EndsWith(string value) => Text(ConditionalFormattingRuleKind.EndsWith, value);
    /// <summary>Formats values that occur more than once in the range.</summary>
    public ExcelConditionalFormatRuleBuilder DuplicateValues() => Add(ConditionalFormattingRuleKind.DuplicateValues);
    /// <summary>Formats values that occur exactly once in the range.</summary>
    public ExcelConditionalFormatRuleBuilder UniqueValues() => Add(ConditionalFormattingRuleKind.UniqueValues);
    /// <summary>Formats blank cells in the range.</summary>
    public ExcelConditionalFormatRuleBuilder Blanks() => Add(ConditionalFormattingRuleKind.Blanks);
    /// <summary>Formats non-blank cells in the range.</summary>
    public ExcelConditionalFormatRuleBuilder NonBlanks() => Add(ConditionalFormattingRuleKind.NonBlanks);

    private ExcelConditionalFormatRuleBuilder Text(ConditionalFormattingRuleKind kind, string value)
    {
        if (string.IsNullOrEmpty(value)) throw new ArgumentException("Conditional formatting text must not be empty.", nameof(value));
        var definition = new ConditionalFormattingRuleDefinition(_range, kind, text: value);
        _sheet.AddConditionalFormat(definition);
        return new ExcelConditionalFormatRuleBuilder(definition);
    }

    private ExcelConditionalFormatRuleBuilder Add(ConditionalFormattingRuleKind kind, params string[] formulas)
    {
        var definition = new ConditionalFormattingRuleDefinition(_range, kind, formulas);
        _sheet.AddConditionalFormat(definition);
        return new ExcelConditionalFormatRuleBuilder(definition);
    }

    private static string Value(object value)
    {
        if (value is null) throw new ArgumentNullException(nameof(value));
        return value switch
        {
            DateTime date => date.ToOADate().ToString(CultureInfo.InvariantCulture),
            string text when !string.IsNullOrEmpty(text) => "\"" + text.Replace("\"", "\"\"") + "\"",
            string => throw new ArgumentException("A conditional formatting value must not be empty.", nameof(value)),
            bool boolean => boolean ? "1" : "0",
            byte number => number.ToString(CultureInfo.InvariantCulture),
            sbyte number => number.ToString(CultureInfo.InvariantCulture),
            short number => number.ToString(CultureInfo.InvariantCulture),
            ushort number => number.ToString(CultureInfo.InvariantCulture),
            int number => number.ToString(CultureInfo.InvariantCulture),
            uint number => number.ToString(CultureInfo.InvariantCulture),
            long number => number.ToString(CultureInfo.InvariantCulture),
            ulong number => number.ToString(CultureInfo.InvariantCulture),
            float number => number.ToString(CultureInfo.InvariantCulture),
            double number => number.ToString(CultureInfo.InvariantCulture),
            decimal number => number.ToString(CultureInfo.InvariantCulture),
            _ => throw new ArgumentException("Conditional formatting values must be numeric, DateTime, Boolean, or text.", nameof(value)),
        };
    }
}

/// <summary>Completes one conditional-formatting rule.</summary>
public sealed class ExcelConditionalFormatRuleBuilder
{
    private readonly ConditionalFormattingRuleDefinition _definition;
    internal ExcelConditionalFormatRuleBuilder(ConditionalFormattingRuleDefinition definition) => _definition = definition;

    /// <summary>Applies a differential style when this rule matches.</summary>
    public ExcelConditionalFormatRuleBuilder Style(Action<ExcelStyleBuilder> configure)
    {
        if (configure is null) throw new ArgumentNullException(nameof(configure));
        var builder = new ExcelStyleBuilder();
        configure(builder);
        var style = builder.Build();
        WorkbookStyleCatalog.ValidateDifferentialStyle(style);
        _definition.SetStyle(style);
        return this;
    }

    /// <summary>Stops later conditional-formatting rules when this rule matches.</summary>
    public ExcelConditionalFormatRuleBuilder StopIfTrue() { _definition.SetStopIfTrue(); return this; }
}
