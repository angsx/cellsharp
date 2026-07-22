using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelAdvancedStylingTests
{
    [Fact]
    public void SchemaFormatsPreserveValuesAndDeduplicateEquivalentFormats()
    {
        var path = TemporaryPath();
        var created = new DateTime(2026, 7, 20, 9, 30, 0, DateTimeKind.Utc);
        var schema = Excel.Schema<ReportRow>()
            .Column(row => row.Amount, column => column.Format("#,##0.00"))
            .Column(row => row.OtherAmount, column => column.Format("#,##0.00"))
            .Column(row => row.Ratio, column => column.Format("0.00%"))
            .Column(row => row.CreatedAt, column => column.Format("dd/MM/yyyy"))
            .Build();

        try
        {
            Excel.Write(path, new[] { new ReportRow(1234.5M, 99.25M, 0.125D, created) }, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var cells = DataRows(document).Single().Elements<Cell>().ToArray();
            var formats = document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
                .Elements<NumberingFormat>()
                .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode!.Value);

            Assert.Equal("1234.5", cells[0].CellValue!.Text);
            Assert.Equal("0.125", cells[2].CellValue!.Text);
            Assert.Equal(created.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), cells[3].CellValue!.Text);
            Assert.Equal("#,##0.00", formats[Format(document, cells[0]).NumberFormatId!.Value]);
            Assert.Equal(Format(document, cells[0]).NumberFormatId!.Value, Format(document, cells[1]).NumberFormatId!.Value);
            Assert.Equal("0.00%", formats[Format(document, cells[2]).NumberFormatId!.Value]);
            Assert.Equal("dd/MM/yyyy", formats[Format(document, cells[3]).NumberFormatId!.Value]);
            Assert.Equal(3, formats.Count);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ExplicitWidthTakesPrecedenceOverAutofitAndInvalidWidthsFailEarly()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<SimpleRow>()
            .Column(row => row.Name, column => column.Width(19.5D))
            .Column(row => row.Count)
            .Build();

        try
        {
            Excel.Write(path, new[] { new SimpleRow("A value deliberately wider than its header", 1) }, schema, options => options.AutoFitColumns());

            using var document = SpreadsheetDocument.Open(path, false);
            var columns = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<Columns>()!
                .Elements<Column>()
                .ToArray();

            Assert.Equal(19.5D, columns[0].Width!.Value);
            Assert.True(columns[1].Width!.Value >= 10D);
        }
        finally
        {
            Delete(path);
        }

        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Schema<SimpleRow>()
            .Column(row => row.Name, column => column.Width(0D)));
    }

    [Fact]
    public void ColumnAlignmentOverridesTheTheme()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<SimpleRow>()
            .Column(row => row.Name, column => column.Align(ExcelHorizontalAlignment.Center))
            .Column(row => row.Count, column => column.Align(ExcelHorizontalAlignment.Right))
            .Build();

        try
        {
            Excel.Write(path, new[] { new SimpleRow("one", 1) }, schema, options => options.Theme(ExcelTheme.Modern));

            using var document = SpreadsheetDocument.Open(path, false);
            var cells = DataRows(document).Single().Elements<Cell>().ToArray();

            Assert.Equal(HorizontalAlignmentValues.Center, Format(document, cells[0]).Alignment!.Horizontal!.Value);
            Assert.Equal(HorizontalAlignmentValues.Right, Format(document, cells[1]).Alignment!.Horizontal!.Value);
        }
        finally
        {
            Delete(path);
        }

        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Schema<SimpleRow>()
            .Column(row => row.Name, column => column.Align((ExcelHorizontalAlignment)42)));
    }

    [Fact]
    public void AlternatingRowsUsesTheActiveTemplateAndCanBeDisabled()
    {
        var alternatingPath = TemporaryPath();
        var plainPath = TemporaryPath();
        var template = new ExcelStyleTemplate(
            dataBackgroundColor: "#FFFFFF",
            alternateRowBackgroundColor: "#F4F7FB");
        var rows = new[] { new SimpleRow("one", 1), new SimpleRow("two", 2) };

        try
        {
            Excel.Write(alternatingPath, rows, options => options.Template(template).AlternatingRows());
            Excel.Write(plainPath, rows, options => options.Template(template));

            using var alternating = SpreadsheetDocument.Open(alternatingPath, false);
            using var plain = SpreadsheetDocument.Open(plainPath, false);
            var alternatingCells = DataRows(alternating).Select(row => row.Elements<Cell>().First()).ToArray();
            var plainCells = DataRows(plain).Select(row => row.Elements<Cell>().First()).ToArray();

            Assert.Equal("FFFFFF", Background(alternating, alternatingCells[0]));
            Assert.Equal("F4F7FB", Background(alternating, alternatingCells[1]));
            Assert.Equal(Background(plain, plainCells[0]), Background(plain, plainCells[1]));
        }
        finally
        {
            Delete(alternatingPath);
            Delete(plainPath);
        }
    }

    [Fact]
    public void CustomTemplateIsReusableImmutableAndSupportsHeaderOverrides()
    {
        var firstPath = TemporaryPath();
        var secondPath = TemporaryPath();
        var template = new ExcelStyleTemplate(
            fontName: "Aptos",
            fontSize: 12D,
            dataTextColor: "#202020",
            dataBackgroundColor: "#FFFFFF",
            headerTextColor: "#FFFFFF",
            headerBackgroundColor: "#1F4E78",
            headerBold: true,
            alternateRowBackgroundColor: "#EAF1FB",
            borderColor: "#B4C7E7");

        try
        {
            Excel.Write(firstPath, new[] { new SimpleRow("one", 1) }, options => options
                .Theme(ExcelTheme.Classic)
                .Template(template)
                .HeaderStyle(style => style.Background("#202020")));
            Excel.Write(secondPath, new[] { new SimpleRow("two", 2) }, options => options.Template(template));

            using var first = SpreadsheetDocument.Open(firstPath, false);
            using var second = SpreadsheetDocument.Open(secondPath, false);
            var firstHeader = HeaderCells(first).First();
            var firstData = DataRows(first).Single().Elements<Cell>().First();
            var secondHeader = HeaderCells(second).First();

            Assert.Equal("202020", Background(first, firstHeader));
            Assert.Equal("FFFFFF", Foreground(first, firstHeader));
            Assert.Equal("Aptos", FontName(first, firstData));
            Assert.Equal("202020", Foreground(first, firstData));
            Assert.Equal("B4C7E7", BorderColor(first, firstData));
            Assert.Equal("1F4E78", Background(second, secondHeader));
            Assert.Equal("#1F4E78", template.HeaderBackgroundColor);
            Assert.Equal("#202020", template.DataTextColor);
        }
        finally
        {
            Delete(firstPath);
            Delete(secondPath);
        }

        Assert.Throws<ArgumentException>(() => new ExcelStyleTemplate(dataTextColor: "blue"));
    }

    private static IEnumerable<Row> DataRows(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.Single().Worksheet!
        .GetFirstChild<SheetData>()!
        .Elements<Row>()
        .Skip(1);

    private static IEnumerable<Cell> HeaderCells(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.Single().Worksheet!
        .GetFirstChild<SheetData>()!
        .Elements<Row>()
        .First()
        .Elements<Cell>();

    private static CellFormat Format(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .CellFormats!
        .Elements<CellFormat>()
        .ElementAt((int)cell.StyleIndex!.Value);

    private static string? Background(SpreadsheetDocument document, Cell cell)
    {
        var format = Format(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fills!
            .Elements<Fill>()
            .ElementAt((int)format.FillId!.Value)
            .PatternFill?
            .ForegroundColor?
            .Rgb?
            .Value?[2..];
    }

    private static string? Foreground(SpreadsheetDocument document, Cell cell)
    {
        var format = Format(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!
            .Elements<Font>()
            .ElementAt((int)format.FontId!.Value)
            .Color?
            .Rgb?
            .Value?[2..];
    }

    private static string? FontName(SpreadsheetDocument document, Cell cell)
    {
        var format = Format(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!
            .Elements<Font>()
            .ElementAt((int)format.FontId!.Value)
            .FontName?
            .Val?
            .Value;
    }

    private static string? BorderColor(SpreadsheetDocument document, Cell cell)
    {
        var format = Format(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Borders!
            .Elements<Border>()
            .ElementAt((int)format.BorderId!.Value)
            .LeftBorder?
            .Color?
            .Rgb?
            .Value?[2..];
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed record ReportRow(decimal Amount, decimal OtherAmount, double Ratio, DateTime CreatedAt);

    private sealed record SimpleRow(string Name, int Count);
}
