using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

namespace CellSharp.Internal;

internal static class XlsxConditionalFormattingWriter
{
    internal static void Apply(Worksheet worksheet, WorkbookStyleCatalog styles, IReadOnlyList<ConditionalFormattingRuleDefinition> definitions)
    {
        for (var index = 0; index < definitions.Count; index++)
        {
            var definition = definitions[index];
            var formatting = new ConditionalFormatting
            {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = definition.Range.ToString() },
            };
            var rule = new ConditionalFormattingRule
            {
                Type = Type(definition.Kind),
                Operator = Operator(definition.Kind),
                Priority = index + 1,
                StopIfTrue = definition.StopIfTrue ? true : null,
                FormatId = definition.Style is null ? null : styles.DifferentialStyleIndex(definition.Style),
                Text = definition.Text,
            };
            foreach (var formula in Formulas(definition)) rule.AppendChild(new Formula(formula));
            formatting.AppendChild(rule);
            worksheet.AppendChild(formatting);
        }
    }

    private static ConditionalFormatValues Type(ConditionalFormattingRuleKind kind) => kind switch
    {
        ConditionalFormattingRuleKind.GreaterThan or ConditionalFormattingRuleKind.GreaterThanOrEqual or
        ConditionalFormattingRuleKind.LessThan or ConditionalFormattingRuleKind.LessThanOrEqual or
        ConditionalFormattingRuleKind.EqualTo or ConditionalFormattingRuleKind.NotEqualTo or
        ConditionalFormattingRuleKind.Between or ConditionalFormattingRuleKind.NotBetween => ConditionalFormatValues.CellIs,
        ConditionalFormattingRuleKind.Formula => ConditionalFormatValues.Expression,
        ConditionalFormattingRuleKind.ContainsText => ConditionalFormatValues.ContainsText,
        ConditionalFormattingRuleKind.BeginsWith => ConditionalFormatValues.BeginsWith,
        ConditionalFormattingRuleKind.EndsWith => ConditionalFormatValues.EndsWith,
        ConditionalFormattingRuleKind.DuplicateValues => ConditionalFormatValues.DuplicateValues,
        ConditionalFormattingRuleKind.UniqueValues => ConditionalFormatValues.UniqueValues,
        ConditionalFormattingRuleKind.Blanks => ConditionalFormatValues.ContainsBlanks,
        ConditionalFormattingRuleKind.NonBlanks => ConditionalFormatValues.NotContainsBlanks,
        _ => throw new ArgumentOutOfRangeException(nameof(kind)),
    };

    private static ConditionalFormattingOperatorValues? Operator(ConditionalFormattingRuleKind kind) => kind switch
    {
        ConditionalFormattingRuleKind.GreaterThan => ConditionalFormattingOperatorValues.GreaterThan,
        ConditionalFormattingRuleKind.GreaterThanOrEqual => ConditionalFormattingOperatorValues.GreaterThanOrEqual,
        ConditionalFormattingRuleKind.LessThan => ConditionalFormattingOperatorValues.LessThan,
        ConditionalFormattingRuleKind.LessThanOrEqual => ConditionalFormattingOperatorValues.LessThanOrEqual,
        ConditionalFormattingRuleKind.EqualTo => ConditionalFormattingOperatorValues.Equal,
        ConditionalFormattingRuleKind.NotEqualTo => ConditionalFormattingOperatorValues.NotEqual,
        ConditionalFormattingRuleKind.Between => ConditionalFormattingOperatorValues.Between,
        ConditionalFormattingRuleKind.NotBetween => ConditionalFormattingOperatorValues.NotBetween,
        ConditionalFormattingRuleKind.ContainsText => ConditionalFormattingOperatorValues.ContainsText,
        ConditionalFormattingRuleKind.BeginsWith => ConditionalFormattingOperatorValues.BeginsWith,
        ConditionalFormattingRuleKind.EndsWith => ConditionalFormattingOperatorValues.EndsWith,
        _ => null,
    };

    private static IEnumerable<string> Formulas(ConditionalFormattingRuleDefinition definition)
    {
        if (definition.Kind is not (ConditionalFormattingRuleKind.ContainsText or ConditionalFormattingRuleKind.BeginsWith or ConditionalFormattingRuleKind.EndsWith)) return definition.Formulas;
        var text = definition.Text ?? throw new InvalidOperationException("A text conditional formatting rule requires text.");
        var escaped = "\"" + text.Replace("\"", "\"\"") + "\"";
        var cell = ExcelRangeReference.ColumnName(definition.Range.FromColumn) + definition.Range.FromRow;
        return definition.Kind switch
        {
            ConditionalFormattingRuleKind.ContainsText => [$"NOT(ISERROR(SEARCH({escaped},{cell})))"],
            ConditionalFormattingRuleKind.BeginsWith => [$"LEFT({cell},LEN({escaped}))={escaped}"],
            ConditionalFormattingRuleKind.EndsWith => [$"RIGHT({cell},LEN({escaped}))={escaped}"],
            _ => throw new ArgumentOutOfRangeException(nameof(definition)),
        };
    }
}
