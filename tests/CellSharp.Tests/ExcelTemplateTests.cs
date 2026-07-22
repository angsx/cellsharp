using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelTemplateTests
{
    private static readonly string[] IdAndNameHeaders = ["Id", "Name"];
    private static readonly string[] ConfiguredHeaders = ["Total", "Birth Date", "Customer Name"];
    private static readonly string[] SharedSchemaHeaders = ["Customer Name", "Id"];

    [Fact]
    public void CreateTemplateCreatesAValidHeaderOnlyWorkbookFromTheSchema()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<CustomerTemplateRow>()
            .Column(row => row.Id)
            .Column(row => row.Name)
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var sheetData = SheetData(document);
            var errors = new OpenXmlValidator().Validate(document).ToArray();

            Assert.Single(document.WorkbookPart!.WorksheetParts);
            Assert.Single(sheetData.Elements<Row>());
            Assert.Equal(IdAndNameHeaders, HeaderCells(document).Select(Text));
            Assert.True(errors.Length == 0, string.Join(Environment.NewLine, errors.Select(error => error.Description)));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void CreateTemplateUsesSchemaOrderOptionalIgnoredColumnsAndColumnPresentation()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<CustomerTemplateRow>()
            .Column(row => row.Total, column => column
                .Header("Total")
                .Format("#,##0.00")
                .Align(ExcelHorizontalAlignment.Right))
            .Column(row => row.BirthDate, column => column
                .Header("Birth Date")
                .Optional()
                .Format("dd/MM/yyyy")
                .Width(16D)
                .Align(ExcelHorizontalAlignment.Center))
            .Column(row => row.Name, column => column.Header("Customer Name"))
            .Column(row => row.InternalCode, column => column.Ignore())
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var columns = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<Columns>()!
                .Elements<Column>()
                .ToArray();
            var formats = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.NumberingFormats!
                .Elements<NumberingFormat>()
                .ToDictionary(format => format.NumberFormatId!.Value, format => format.FormatCode!.Value);

            Assert.Equal(ConfiguredHeaders, HeaderCells(document).Select(Text));
            Assert.Equal(3, columns.Length);
            Assert.Equal(16D, columns[1].Width!.Value);
            Assert.Equal("#,##0.00", formats[ColumnFormat(document, columns[0]).NumberFormatId!.Value]);
            Assert.Equal("dd/MM/yyyy", formats[ColumnFormat(document, columns[1]).NumberFormatId!.Value]);
            Assert.Equal(HorizontalAlignmentValues.Right, ColumnFormat(document, columns[0]).Alignment!.Horizontal!.Value);
            Assert.Equal(HorizontalAlignmentValues.Center, ColumnFormat(document, columns[1]).Alignment!.Horizontal!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void CreateTemplateSupportsThemesTemplatesHeaderOverridesAutofitAndFrozenHeaders()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<CustomerTemplateRow>()
            .Column(row => row.Id, column => column.Width(18D))
            .Column(row => row.Name, column => column.Header("A long customer heading"))
            .Build();
        var template = new ExcelStyleTemplate(
            fontName: "Aptos",
            dataTextColor: "#202020",
            headerTextColor: "#FFFFFF",
            headerBackgroundColor: "#1F4E78",
            borderColor: "#B4C7E7");

        try
        {
            Excel.CreateTemplate(path, schema, options => options
                .Theme(ExcelTheme.Classic)
                .Template(template)
                .HeaderStyle(style => style.Background("#202020"))
                .AutoFitColumns()
                .AlternatingRows()
                .FreezeHeaderRow());

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single().Worksheet!;
            var columns = worksheet.GetFirstChild<Columns>()!.Elements<Column>().ToArray();
            var header = HeaderCells(document).First();
            var pane = worksheet.GetFirstChild<SheetViews>()!.Elements<SheetView>().Single().Pane!;

            Assert.Equal(18D, columns[0].Width!.Value);
            Assert.Equal("202020", Background(document, header));
            Assert.Equal("FFFFFF", Foreground(document, header));
            Assert.Equal("Aptos", FontName(document, header));
            Assert.True(columns[1].Width!.Value >= "A long customer heading".Length + 2D);
            Assert.Equal(PaneStateValues.Frozen, pane.State!.Value);
            Assert.Single(SheetData(document).Elements<Row>());
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void CreateTemplateSupportsBuiltInThemes()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<CustomerTemplateRow>()
            .Column(row => row.Name)
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema, options => options.Theme(ExcelTheme.Modern));

            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Equal("1F4E78", Background(document, HeaderCells(document).First()));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void CreateTemplateOverwritesAndSharesTheSchemaPipelineWithWrite()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<CustomerTemplateRow>()
            .Column(row => row.Name, column => column.Header("Customer Name"))
            .Column(row => row.Id)
            .Build();

        try
        {
            Excel.Write(path, new[] { new CustomerTemplateRow { Id = 7, Name = "Ada" } }, schema);
            Excel.CreateTemplate(path, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            Assert.Equal(SharedSchemaHeaders, HeaderCells(document).Select(Text));
            Assert.Single(SheetData(document).Elements<Row>());
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ACustomStyleTemplateCanBeReusedAcrossGeneratedTemplates()
    {
        var firstPath = TemporaryPath();
        var secondPath = TemporaryPath();
        var schema = Excel.Schema<CustomerTemplateRow>()
            .Column(row => row.Name)
            .Build();
        var template = new ExcelStyleTemplate(
            fontName: "Arial",
            dataTextColor: "#202020",
            headerTextColor: "#FFFFFF",
            headerBackgroundColor: "#1F4E78");

        try
        {
            Excel.CreateTemplate(firstPath, schema, options => options.Template(template));
            Excel.CreateTemplate(secondPath, schema, options => options.Template(template));

            using var first = SpreadsheetDocument.Open(firstPath, false);
            using var second = SpreadsheetDocument.Open(secondPath, false);

            Assert.Equal("1F4E78", Background(first, HeaderCells(first).First()));
            Assert.Equal("1F4E78", Background(second, HeaderCells(second).First()));
            Assert.Equal("Arial", FontName(first, HeaderCells(first).First()));
            Assert.Equal("#1F4E78", template.HeaderBackgroundColor);
        }
        finally
        {
            Delete(firstPath);
            Delete(secondPath);
        }
    }

    private static SheetData SheetData(SpreadsheetDocument document) => document.WorkbookPart!.WorksheetParts.Single().Worksheet!
        .GetFirstChild<SheetData>()!;

    private static IEnumerable<Cell> HeaderCells(SpreadsheetDocument document) => SheetData(document)
        .Elements<Row>()
        .First()
        .Elements<Cell>();

    private static CellFormat ColumnFormat(SpreadsheetDocument document, Column column) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .CellFormats!
        .Elements<CellFormat>()
        .ElementAt((int)column.Style!.Value);

    private static string? Background(SpreadsheetDocument document, Cell cell)
    {
        var format = CellFormat(document, cell);
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
        var format = CellFormat(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!
            .Elements<Font>()
            .ElementAt((int)format.FontId!.Value)
            .Color?
            .Rgb?
            .Value?[2..];
    }

    private static string? FontName(SpreadsheetDocument document, Cell cell)
    {
        var format = CellFormat(document, cell);
        return document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!.Fonts!
            .Elements<Font>()
            .ElementAt((int)format.FontId!.Value)
            .FontName?
            .Val?
            .Value;
    }

    private static CellFormat CellFormat(SpreadsheetDocument document, Cell cell) => document.WorkbookPart!.WorkbookStylesPart!.Stylesheet!
        .CellFormats!
        .Elements<CellFormat>()
        .ElementAt((int)cell.StyleIndex!.Value);

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-template-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class CustomerTemplateRow
    {
        public int Id { get; set; }

        public string? Name { get; set; }

        public DateTime BirthDate { get; set; }

        public decimal Total { get; set; }

        public string? InternalCode { get; set; }
    }
}
