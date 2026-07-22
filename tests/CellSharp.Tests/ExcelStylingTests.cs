using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelStylingTests
{
    private static readonly string[] CustomHeaders = ["Customer Name", "Customer ID"];

    [Fact]
    public void WriteAppliesTheDefaultTheme()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, Customers());

            using var document = SpreadsheetDocument.Open(path, false);
            var cells = Cells(document).ToArray();

            Assert.NotNull(document.WorkbookPart!.WorkbookStylesPart);
            Assert.NotEqual(0U, cells[0].StyleIndex!.Value);
            Assert.NotEqual(cells[0].StyleIndex!.Value, cells[2].StyleIndex!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Theory]
    [InlineData(ExcelTheme.Default)]
    [InlineData(ExcelTheme.Modern)]
    [InlineData(ExcelTheme.Classic)]
    [InlineData(ExcelTheme.Minimal)]
    public void WriteSupportsEveryBuiltInTheme(ExcelTheme theme)
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, Customers(), options => options.Theme(theme));

            using var document = SpreadsheetDocument.Open(path, false);
            var header = Cells(document).First();

            Assert.NotEqual(0U, header.StyleIndex!.Value);
            Assert.NotNull(HeaderBackground(document, header));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteAppliesCustomHeaderStylesOverTheTheme()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, Customers(), options => options
                .Theme(ExcelTheme.Modern)
                .HeaderStyle(style => style.Background("#202020")));

            using var document = SpreadsheetDocument.Open(path, false);
            var header = Cells(document).First();

            Assert.Equal("202020", HeaderBackground(document, header));
            Assert.Equal("FFFFFF", HeaderForeground(document, header));
            Assert.True(HeaderIsBold(document, header));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteRejectsInvalidHeaderColorsBeforeCreatingAWorkbook()
    {
        var path = TemporaryPath();

        var exception = Assert.Throws<ArgumentException>(() => Excel.Write(path, Customers(), options => options
            .HeaderStyle(style => style.Background("blue"))));

        Assert.Contains("#RRGGBB", exception.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void WriteKeepsValuesAndFormulaLikeTextUnchangedWhenStyled()
    {
        var path = TemporaryPath();
        var customer = new StyledCustomer { Id = 7, Name = "=SUM(A1:A2)" };

        try
        {
            Excel.Write(path, new[] { customer }, options => options.Theme(ExcelTheme.Classic));

            var result = Excel.Read<StyledCustomer>(path);
            using var document = SpreadsheetDocument.Open(path, false);
            var valueCell = Cells(document).Skip(3).First();

            Assert.True(result.IsValid);
            Assert.Equal(customer.Name, Assert.Single(result.Rows).Name);
            Assert.Equal(DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString, valueCell.DataType!.Value);
            Assert.Null(valueCell.CellFormula);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteReusesStyleRecordsAcrossDataCells()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, Customers(), options => options.Theme(ExcelTheme.Modern));

            using var document = SpreadsheetDocument.Open(path, false);
            var rows = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .ToArray();
            var dataStyleIndexes = rows.Skip(1).SelectMany(row => row.Elements<Cell>()).Select(cell => cell.StyleIndex!.Value).Distinct().ToArray();
            var formats = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!.Elements<CellFormat>().ToArray();

            Assert.Single(dataStyleIndexes);
            Assert.InRange(formats.Length, 1, 4);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteSupportsSchemasThemesWidthEstimationAndFrozenHeaders()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<StyledCustomer>()
            .Column(customer => customer.Name, column => column.Header("Customer Name"))
            .Column(customer => customer.Id, column => column.Header("Customer ID"))
            .Build();

        try
        {
            Excel.Write(path, Customers(), schema, options => options
                .Theme(ExcelTheme.Modern)
                .HeaderStyle(style => style.Bold().Foreground("#FFFFFF"))
                .AutoFitColumns()
                .FreezeHeaderRow());

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var headers = worksheet.GetFirstChild<SheetData>()!.Elements<Row>().First().Elements<Cell>().Select(Text).ToArray();
            var columns = worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
            var pane = worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().Single().Pane!;

            Assert.Equal(CustomHeaders, headers);
            Assert.Equal(2, columns.Length);
            Assert.All(columns, column => Assert.True(column.CustomWidth!.Value));
            Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    private static IEnumerable<StyledCustomer> Customers() =>
    [
        new StyledCustomer { Id = 1, Name = "Alice" },
        new StyledCustomer { Id = 2, Name = "Bob" },
    ];

    private static IEnumerable<Cell> Cells(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.Single().Worksheet!
        .GetFirstChild<SheetData>()!
        .Elements<Row>()
        .SelectMany(row => row.Elements<Cell>());

    private static string? HeaderBackground(SpreadsheetDocument document, Cell cell)
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

    private static string? HeaderForeground(SpreadsheetDocument document, Cell cell)
    {
        var format = Format(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!
            .Elements<Font>()
            .ElementAt((int)format.FontId!.Value)
            .Color?
            .Rgb?
            .Value?[2..];
    }

    private static bool HeaderIsBold(SpreadsheetDocument document, Cell cell)
    {
        var format = Format(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!
            .Elements<Font>()
            .ElementAt((int)format.FontId!.Value)
            .Elements<Bold>()
            .Any();
    }

    private static CellFormat Format(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .CellFormats!
        .Elements<CellFormat>()
        .ElementAt((int)cell.StyleIndex!.Value);

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class StyledCustomer
    {
        public int Id { get; set; }

        public string? Name { get; set; }
    }
}
