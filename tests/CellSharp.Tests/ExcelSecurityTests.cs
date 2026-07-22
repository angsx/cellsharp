using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelSecurityTests
{
    private static readonly string[] FormulaLookingValues = ["=SUM(A1:A2)", "+CMD()", "-CMD()", "@SUM(A1:A2)", "\t=SUM(A1:A2)", "\r\n=SUM(A1:A2)", " =SUM(A1:A2)", "'=SUM(A1:A2)"];

    [Fact]
    public void WriteKeepsFormulaLookingUserStringsAsLiteralInlineText()
    {
        var path = TemporaryPath();
        var serializedValues = FormulaLookingValues.Select(value => value.Replace("\r\n", "\n", StringComparison.Ordinal)).ToArray();

        try
        {
            Excel.Write(path, FormulaLookingValues.Select(value => new TextRow { Value = value }));

            using var document = SpreadsheetDocument.Open(path, false);
            var cells = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Select(row => Assert.Single(row.Elements<Cell>()))
                .ToArray();

            Assert.Equal(serializedValues, cells.Select(Text));
            Assert.All(cells, cell =>
            {
                Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
                Assert.Null(cell.CellFormula);
            });
            Assert.Equal(serializedValues, Excel.Read<TextRow>(path).Rows.Select(row => row.Value));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ConverterOutputThatLooksLikeAFormulaRemainsLiteralText()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<DomainRow>()
            .Column(row => row.Value, column => column.ConvertWith(FormulaLookingConverter.Instance))
            .Build();

        try
        {
            Excel.Write(path, [new DomainRow { Value = new DomainValue("=2+2") }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var cell = document.WorkbookPart!.WorksheetParts.Single().Worksheet!.GetFirstChild<SheetData>()!
                .Elements<Row>()
                .Skip(1)
                .Single()
                .Elements<Cell>()
                .Single();

            Assert.Equal("=2+2", Text(cell));
            Assert.Equal(CellValues.InlineString, cell.DataType!.Value);
            Assert.Null(cell.CellFormula);
            Assert.Equal("=2+2", Assert.Single(Excel.Read<DomainRow>(path, schema).Rows).Value!.Text);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void TemplateHeadersAndAllowedValuesRemainLiteralText()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<TextRow>()
            .Column(row => row.Value, column => column
                .Header("=Header")
                .AllowedValues(FormulaLookingValues))
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var cells = document.WorkbookPart!.WorksheetParts
                .SelectMany(part => part.Worksheet!.GetFirstChild<SheetData>()?.Elements<Row>() ?? Enumerable.Empty<Row>())
                .SelectMany(row => row.Elements<Cell>())
                .ToArray();

            Assert.Contains(cells, cell => Text(cell) == "=Header");
            Assert.All(cells, cell => Assert.Null(cell.CellFormula));
            Assert.All(cells.Where(cell => Text(cell).Length > 0), cell => Assert.Equal(CellValues.InlineString, cell.DataType!.Value));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadRejectsPackagesAndPartsBeyondConfiguredLimits()
    {
        using var oversizedPackage = new MemoryStream(new byte[101]);
        var packageOptions = new ExcelReadOptions(maxPackageBytes: 100);
        Assert.Throws<InvalidDataException>(() => Excel.Read<TextRow>(oversizedPackage, packageOptions));
        using var oversizedWorkbook = new MemoryStream(new byte[101]);
        Assert.Throws<InvalidDataException>(() => Excel.Open(oversizedWorkbook, packageOptions));

        var path = TemporaryPath();
        try
        {
            CreateWorkbook(path,
                Row(TextCell("A1", "Value")),
                Row(TextCell("A2", new string('x', 4096))));

            var partOptions = new ExcelReadOptions(maxCharactersInPart: 1024);
            var exception = Assert.Throws<System.Xml.XmlException>(() => Excel.Read<TextRow>(path, partOptions));
            Assert.Contains("MaxCharactersInDocument", exception.Message, StringComparison.Ordinal);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadRejectsExcessiveRowsErrorsAndInvalidPhysicalColumns()
    {
        var rowPath = TemporaryPath();
        var errorPath = TemporaryPath();
        var columnPath = TemporaryPath();
        var duplicatePath = TemporaryPath();
        try
        {
            CreateWorkbook(rowPath,
                Row(TextCell("A1", "Value")),
                Row(TextCell("A2", "one")),
                Row(TextCell("A3", "two")));
            Assert.Throws<InvalidDataException>(() => Excel.Read<TextRow>(rowPath, new ExcelReadOptions(maxRows: 2)));

            CreateWorkbook(errorPath,
                Row(TextCell("A1", "Number")),
                Row(TextCell("A2", "bad")),
                Row(TextCell("A3", "also bad")));
            Assert.Throws<InvalidDataException>(() => Excel.Read<NumberRow>(errorPath, new ExcelReadOptions(maxErrors: 1)));

            CreateWorkbook(columnPath, Row(TextCell("XFE1", "Value")));
            Assert.Throws<InvalidDataException>(() => Excel.Read<TextRow>(columnPath));

            CreateWorkbook(duplicatePath, Row(TextCell("A1", "Value"), TextCell("A1", "Duplicate")));
            Assert.Throws<InvalidDataException>(() => Excel.Read<TextRow>(duplicatePath));
        }
        finally
        {
            Delete(rowPath);
            Delete(errorPath);
            Delete(columnPath);
            Delete(duplicatePath);
        }
    }

    [Fact]
    public void ReadRejectsMalformedOrConflictingPhysicalCellCoordinates()
    {
        var invalidRowPath = TemporaryPath();
        var mismatchedRowPath = TemporaryPath();
        try
        {
            var invalidRow = new Row { RowIndex = 1U };
            invalidRow.AppendChild(TextCell("A0", "Value"));
            CreateWorkbook(invalidRowPath, invalidRow);
            Assert.Throws<InvalidDataException>(() => Excel.Read<TextRow>(invalidRowPath));

            var mismatchedRow = new Row { RowIndex = 2U };
            mismatchedRow.AppendChild(TextCell("A3", "Value"));
            CreateWorkbook(mismatchedRowPath, mismatchedRow);
            Assert.Throws<InvalidDataException>(() => Excel.Read<TextRow>(mismatchedRowPath));
        }
        finally
        {
            Delete(invalidRowPath);
            Delete(mismatchedRowPath);
        }
    }

    [Fact]
    public void HyperlinksRequireHttpHttpsOrAnInternalLocation()
    {
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Links", sheet => sheet.Cell("A1").Hyperlink("file:///tmp/private")));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Links", sheet => sheet.Cell("A1").Hyperlink("javascript:alert(1)")));
        Assert.Throws<ArgumentException>(() => Excel.Workbook().AddSheet("Links", sheet => sheet.Cell("A1").Hyperlink("#")));

        Excel.Workbook().AddSheet("Links", sheet =>
        {
            sheet.Cell("A1").Hyperlink("https://example.com");
            sheet.Cell("A2").Hyperlink("#Links!A1");
        });
    }

    [Fact]
    public void ReadRejectsMissingOrInvalidSharedStringReferences()
    {
        var path = TemporaryPath();
        try
        {
            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData(
                    Row(TextCell("A1", "Value")),
                    Row(new Cell { CellReference = "A2", DataType = CellValues.SharedString, CellValue = new CellValue("99") })));
                var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
                worksheetPart.Worksheet.Save();
                workbookPart.Workbook.Save();
            }

            Assert.Throws<InvalidDataException>(() => Excel.Read<TextRow>(path));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadOptionsRejectNonPositiveSecurityLimits()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelReadOptions(maxRows: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelReadOptions(maxErrors: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelReadOptions(maxPackageBytes: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelReadOptions(maxCharactersInPart: 0));
    }

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static void CreateWorkbook(string path, params Row[] rows)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(new SheetData(rows));
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = 1U,
            Name = "Sheet1",
        });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static Row Row(params Cell[] cells) => new(cells);

    private static Cell TextCell(string reference, string value) => new()
    {
        CellReference = reference,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(value)),
    };

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-security-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class TextRow
    {
        public string? Value { get; set; }
    }

    private sealed class NumberRow
    {
        public int Number { get; set; }
    }

    private sealed class DomainRow
    {
        public DomainValue? Value { get; set; }
    }

    private sealed class DomainValue
    {
        internal DomainValue(string text) => Text = text;

        internal string Text { get; }
    }

    private sealed class FormulaLookingConverter : IExcelValueConverter<DomainValue?, string>
    {
        internal static FormulaLookingConverter Instance { get; } = new();

        public string Write(DomainValue? value) => value?.Text ?? string.Empty;

        public bool TryRead(string value, out DomainValue? converted)
        {
            converted = new DomainValue(value);
            return true;
        }
    }
}
