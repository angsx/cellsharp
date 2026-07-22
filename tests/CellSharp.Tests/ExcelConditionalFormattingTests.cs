using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelConditionalFormattingTests
{
    private static readonly int[] Priorities = [1, 2, 3, 4, 5, 6, 7, 8];
    private static readonly string[] BetweenFormulas = ["1", "3"];
    [Fact]
    public void ValueComparisonRulesRenderNativeRulesInDeclarationOrder()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Rules", sheet =>
            {
                sheet.Range("A1:A10").ConditionalFormat().GreaterThan(10).Style(s => s.FillColor("#C6EFCE"));
                sheet.Range("A1:A10").ConditionalFormat().GreaterThanOrEqual(9);
                sheet.Range("A1:A10").ConditionalFormat().LessThan(0);
                sheet.Range("A1:A10").ConditionalFormat().LessThanOrEqual(-1);
                sheet.Range("A1:A10").ConditionalFormat().EqualTo(5);
                sheet.Range("A1:A10").ConditionalFormat().NotEqualTo(6);
                sheet.Range("A1:A10").ConditionalFormat().Between(1, 3);
                sheet.Range("A1:A10").ConditionalFormat().NotBetween(4, 8);
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var rules = Rules(document).ToArray();
            Assert.Equal(8, rules.Length);
            Assert.All(rules, rule => Assert.Equal(ConditionalFormatValues.CellIs, rule.Type!.Value));
            Assert.Equal(Priorities, rules.Select(rule => rule.Priority!.Value));
            Assert.Equal(BetweenFormulas, rules[6].Elements<Formula>().Select(formula => formula.Text));
            AssertValid(document);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void FormulaTextBlankDuplicateAndUniqueRulesRenderNatively()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Rules", sheet =>
            {
                sheet.Range("A2:A100").ConditionalFormat().Formula("=A2<TODAY()").Style(s => s.FontColor("#9C0006"));
                sheet.Range("B2:B100").ConditionalFormat().ContainsText("ERROR");
                sheet.Range("C2:C100").ConditionalFormat().BeginsWith("WARN");
                sheet.Range("D2:D100").ConditionalFormat().EndsWith("!");
                sheet.Range("E2:E100").ConditionalFormat().DuplicateValues();
                sheet.Range("F2:F100").ConditionalFormat().UniqueValues();
                sheet.Range("G2:G100").ConditionalFormat().Blanks();
                sheet.Range("H2:H100").ConditionalFormat().NonBlanks().StopIfTrue();
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var rules = Rules(document).ToArray();
            Assert.Equal("A2<TODAY()", rules[0].GetFirstChild<Formula>()!.Text);
            Assert.Equal("ERROR", rules[1].Text!.Value);
            Assert.Contains("SEARCH(\"ERROR\",B2)", rules[1].GetFirstChild<Formula>()!.Text);
            Assert.Equal(ConditionalFormatValues.DuplicateValues, rules[4].Type!.Value);
            Assert.Equal(ConditionalFormatValues.UniqueValues, rules[5].Type!.Value);
            Assert.Equal(ConditionalFormatValues.ContainsBlanks, rules[6].Type!.Value);
            Assert.Equal(ConditionalFormatValues.NotContainsBlanks, rules[7].Type!.Value);
            Assert.True(rules[7].StopIfTrue!.Value);
            AssertValid(document);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void DifferentialStylesAreDeduplicatedAndCannotUseUnsupportedProperties()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Rules", sheet =>
            {
                sheet.Range("A1:A10").ConditionalFormat().GreaterThan(10).Style(s => s.Bold().FontColor("#006100").FillColor("#C6EFCE").Border(b => b.Bottom(ExcelBorderStyle.Thin, "#006100")).NumberFormat("0.00%"));
                sheet.Range("B1:B10").ConditionalFormat().LessThan(0).Style(s => s.Bold().FontColor("#006100").FillColor("#C6EFCE").Border(b => b.Bottom(ExcelBorderStyle.Thin, "#006100")).NumberFormat("0.00%"));
            }).Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var styles = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!;
            Assert.Single(styles.DifferentialFormats!.Elements<DifferentialFormat>());
            AssertValid(document);
        }
        finally { Delete(path); }

        var invalidPath = TemporaryPath();
        try
        {
            Assert.Throws<InvalidOperationException>(() => Excel.Workbook().AddSheet("Rules", sheet =>
                sheet.Range("A1").ConditionalFormat().EqualTo(1).Style(s => s.FontSize(12))).Write(invalidPath));
        }
        finally { Delete(invalidPath); }
    }

    [Fact]
    public void LargeColumnRangeDoesNotMaterializeCells()
    {
        var path = TemporaryPath();
        try
        {
            Excel.Workbook().AddSheet("Sparse", sheet => sheet.Range("A:A").ConditionalFormat().GreaterThan(0)).Write(path);
            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            Assert.Empty(worksheet.GetFirstChild<SheetData>()!.Descendants<Cell>());
            Assert.Equal("A1:A1048576", worksheet.GetFirstChild<ConditionalFormatting>()!.SequenceOfReferences!.InnerText);
            AssertValid(document);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void ConditionalFormattingCoexistsWithOffsetTableValidationMergeAndMultipleSheets()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Row>()
            .SheetName("Orders")
            .AsTable()
            .Column(row => row.Id)
            .Column(row => row.Status, column => column.AllowedValues("OK", "ERROR"))
            .Column(row => row.Amount)
            .Build();
        try
        {
            Excel.Workbook()
                .AddSheet([new Row(1, "OK", 20), new Row(2, "ERROR", -1)], schema, sheet =>
                {
                    sheet.DataStartAt("B4");
                    sheet.Range("D5:D100").ConditionalFormat().LessThan(0).Style(s => s.FillColor("#FFC7CE"));
                    sheet.Merge("B1:D1").Value("Orders");
                })
                .AddSheet("Summary", sheet => sheet.Range("A1:A10").ConditionalFormat().DuplicateValues())
                .Write(path);

            using var document = SpreadsheetDocument.Open(path, false);
            var orders = document.WorkbookPart!.WorksheetParts.First().Worksheet!;
            Assert.Equal("D5:D100", orders.GetFirstChild<ConditionalFormatting>()!.SequenceOfReferences!.InnerText);
            Assert.NotNull(orders.GetFirstChild<DataValidations>());
            Assert.NotNull(orders.GetFirstChild<MergeCells>());
            Assert.Single(document.WorkbookPart.WorksheetParts.First().TableDefinitionParts);
            Assert.Equal(3, document.WorkbookPart.WorksheetParts.Count()); // Orders, Summary, hidden validation list
            AssertValid(document);
        }
        finally { Delete(path); }
    }

    [Fact]
    public void InvalidFormulaAndRepeatedRuleConfigurationFailEarly()
    {
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Rules", sheet => sheet.Cell("A1").ConditionalFormat().Formula(" ")));
        var builder = Excel.Workbook().AddSheet("Rules", sheet =>
        {
            var rule = sheet.Cell("A1").ConditionalFormat().GreaterThan(1);
            rule.StopIfTrue();
            Assert.Throws<InvalidOperationException>(() => rule.StopIfTrue());
        });
        Assert.NotNull(builder);
    }

    private static IEnumerable<ConditionalFormattingRule> Rules(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.Single().Worksheet!.Elements<ConditionalFormatting>().SelectMany(formatting => formatting.Elements<ConditionalFormattingRule>());
    private static void AssertValid(SpreadsheetDocument document)
    {
        var errors = new OpenXmlValidator().Validate(document).ToArray();
        Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
    }
    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-conditional-{Guid.NewGuid():N}.xlsx");
    private static void Delete(string path) { if (File.Exists(path)) File.Delete(path); }
    private sealed record Row(int Id, string Status, decimal Amount);
}
