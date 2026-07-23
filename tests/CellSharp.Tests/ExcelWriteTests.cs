using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelWriteTests
{
    private static readonly string[] CustomerHeaders = ["Id", "Name", "IsActive", "CreatedAt", "ExternalId", "Balance"];

    [Fact]
    public void WriteCreatesWorkbookWithHeadersAndScalarValues()
    {
        var path = TemporaryPath();
        var createdAt = new DateTime(2026, 7, 20, 9, 30, 0, DateTimeKind.Utc);
        var customers = new[]
        {
            new Customer(7, "Ada", true, createdAt, Guid.Parse("8b7f9a19-6520-4f29-9bb8-50b763162e7d"), null),
        };

        try
        {
            Excel.Write(path, customers);

            using var document = SpreadsheetDocument.Open(path, false);
            var worksheet = document.WorkbookPart!.WorksheetParts.Single();
            var rows = worksheet.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().ToArray();

            Assert.Equal(2, rows.Length);
            Assert.Equal(CustomerHeaders, rows[0].Elements<Cell>().Select(Text));

            var values = rows[1].Elements<Cell>().ToArray();
            Assert.Equal("7", values[0].CellValue!.Text);
            Assert.Equal("Ada", Text(values[1]));
            Assert.Equal("1", values[2].CellValue!.Text);
            Assert.Equal(createdAt.ToOADate().ToString(System.Globalization.CultureInfo.InvariantCulture), values[3].CellValue!.Text);
            Assert.NotNull(values[3].StyleIndex);
            Assert.NotEqual(0U, values[3].StyleIndex!.Value);
            Assert.Equal("8b7f9a19-6520-4f29-9bb8-50b763162e7d", Text(values[4]));
            Assert.Null(values[5].CellValue);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteCreatesHeadersForAnEmptySequence()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, Array.Empty<Customer>());

            using var document = SpreadsheetDocument.Open(path, false);
            var rows = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .ToArray();

            Assert.Single(rows);
            Assert.Equal(CustomerHeaders, rows[0].Elements<Cell>().Select(Text));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteUsesTwoDecimalPlacesForDecimalValuesByDefault()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, [new DecimalValue(1M)]);

            using var document = SpreadsheetDocument.Open(path, false);
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Single()
                .Elements<Cell>()
                .Single();
            var format = document.WorkbookPart.WorkbookStylesPart!.Stylesheet!.CellFormats!
                .Elements<CellFormat>()
                .ElementAt((int)cell.StyleIndex!.Value);
            var numberFormat = document.WorkbookPart.WorkbookStylesPart.Stylesheet.NumberingFormats!
                .Elements<NumberingFormat>()
                .Single(candidate => candidate.NumberFormatId!.Value == format.NumberFormatId!.Value);

            Assert.Equal("0.00", numberFormat.FormatCode!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteKeepsFormulaLikeTextAsText()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, new[] { new FormulaCandidate("=1+1") });

            using var document = SpreadsheetDocument.Open(path, false);
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Single()
                .Elements<Cell>()
                .Single();

            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Equal("=1+1", Text(cell));
            Assert.Null(cell.CellFormula);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteReplacesAnExistingFile()
    {
        var path = TemporaryPath();

        try
        {
            Excel.Write(path, new[] { new FormulaCandidate("first") });
            Excel.Write(path, new[] { new FormulaCandidate("second") });

            using var document = SpreadsheetDocument.Open(path, false);
            var value = document.WorkbookPart!.WorksheetParts.Single().Worksheet!
                .GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Single()
                .Elements<Cell>()
                .Single();

            Assert.Equal("second", Text(value));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteRejectsATypeWithoutExportableProperties()
    {
        var path = TemporaryPath();

        var exception = Assert.Throws<InvalidOperationException>(() => Excel.Write(path, Array.Empty<Empty>()));

        Assert.Contains("no public readable properties", exception.Message);
        Assert.False(File.Exists(path));
    }

    [Fact]
    public void WriteRejectsANonexistentDirectory()
    {
        var path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N"), "customers.xlsx");

        Assert.Throws<DirectoryNotFoundException>(() => Excel.Write(path, Array.Empty<Customer>()));
    }

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed record Customer(int Id, string Name, bool IsActive, DateTime CreatedAt, Guid ExternalId, decimal? Balance);

    private sealed record DecimalValue(decimal Value);

    private sealed record FormulaCandidate(string Value);

    private sealed class Empty;
}
