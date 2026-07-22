using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelImportResilienceTests
{
    [Fact]
    public void ReadSkipsLeadingAndStyledEmptyRowsWithoutTrustingWorksheetDimension()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(
                path,
                [
                    new Row { RowIndex = 1U, CustomFormat = true, StyleIndex = 0U },
                    Row(2U, TextCell("A2", "Name")),
                    Row(3U, new Cell { CellReference = "A3", StyleIndex = 0U }),
                    Row(4U, TextCell("A4", "Ada")),
                    Row(500000U, new Cell { CellReference = "A500000", StyleIndex = 0U }),
                ],
                "A1:XFD1048576");

            var result = Excel.Read<TextRow>(path);

            Assert.True(result.IsValid);
            Assert.Equal("Ada", Assert.Single(result.Rows).Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadHandlesSparseRowsAndPreservesRowsThatAreFullyValid()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path,
            [
                Row(1U, TextCell("A1", "Name"), TextCell("B1", "Age"), TextCell("D1", "Id")),
                Row(2U, TextCell("A2", "Ada"), NumberCell("D2", "7")),
                Row(5U, TextCell("A5", "Grace"), NumberCell("B5", "30"), NumberCell("D5", "8")),
            ]);

            var result = Excel.Read<SparseRow>(path);

            var row = Assert.Single(result.Rows);
            Assert.Equal("Grace", row.Name);
            Assert.Equal(30, row.Age);
            Assert.Equal(8, row.Id);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.RequiredValueMissing, error.Code);
            Assert.Equal("B2", error.CellReference);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadPreservesEmptyInlineAndSharedStringsInOtherwiseSignificantRows()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(
                path,
                [
                    Row(1U, TextCell("A1", "Name"), TextCell("B1", "Value")),
                    Row(2U, TextCell("A2", "Ada"), TextCell("B2", string.Empty)),
                    Row(3U, TextCell("A3", "Grace"), new Cell
                    {
                        CellReference = "B3",
                        DataType = CellValues.SharedString,
                        CellValue = new CellValue("0"),
                    }),
                ],
                sharedStrings: new SharedStringTable(new SharedStringItem(new Text(string.Empty))));

            var result = Excel.Read<StringPair>(path);

            Assert.True(result.IsValid);
            Assert.Equal(["Ada", "Grace"], result.Rows.Select(row => row.Name));
            Assert.All(result.Rows, row => Assert.Equal(string.Empty, row.Value));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadUsesFormulaCachesAndReportsFormulaWithoutCacheAndExcelErrors()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path,
            [
                Row(1U, TextCell("A1", "Amount")),
                Row(2U, new Cell
                {
                    CellReference = "A2",
                    CellFormula = new CellFormula("1+1"),
                    CellValue = new CellValue("2"),
                }),
                Row(3U, new Cell { CellReference = "A3", CellFormula = new CellFormula("1+1") }),
                Row(4U, new Cell
                {
                    CellReference = "A4",
                    DataType = CellValues.Error,
                    CellValue = new CellValue("#DIV/0!"),
                }),
                Row(5U, NumberCell("A5", "4")),
            ]);

            var result = Excel.Read<FormulaRow>(path);

            Assert.Equal([2, 4], result.Rows.Select(row => row.Amount));
            Assert.Equal(2, result.Errors.Count);
            Assert.All(result.Errors, error => Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code));
            Assert.Contains(result.Errors, error => error.Message == "Formula cells require a cached value to be imported.");
            Assert.Contains(result.Errors, error => error.Message.Contains("Excel error value '#DIV/0!'", StringComparison.Ordinal));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadPreservesWhitespaceTextAndReportsWhitespaceOnlyNumericValues()
    {
        var path = TemporaryPath();

        try
        {
            CreateWorkbook(path,
            [
                Row(1U, TextCell("A1", "Text"), TextCell("B1", "Amount")),
                Row(2U, TextCell("A2", "  "), TextCell("B2", "  ")),
                Row(3U, TextCell("A3", "kept"), NumberCell("B3", "12")),
            ]);

            var result = Excel.Read<WhitespaceRow>(path);

            Assert.Equal("kept", Assert.Single(result.Rows).Text);
            var error = Assert.Single(result.Errors);
            Assert.Equal(ExcelReadErrorCode.InvalidValue, error.Code);
            Assert.Equal("B2", error.CellReference);
        }
        finally
        {
            Delete(path);
        }
    }

    private static void CreateWorkbook(
        string path,
        IEnumerable<Row> rows,
        string? dimension = null,
        SharedStringTable? sharedStrings = null)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        if (sharedStrings is not null)
        {
            var sharedStringPart = workbookPart.AddNewPart<SharedStringTablePart>();
            sharedStringPart.SharedStringTable = sharedStrings;
            sharedStringPart.SharedStringTable.Save();
        }

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var sheetData = new SheetData(rows);
        worksheetPart.Worksheet = dimension is null
            ? new Worksheet(sheetData)
            : new Worksheet(new SheetDimension { Reference = dimension }, sheetData);
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static Row Row(uint rowIndex, params Cell[] cells)
    {
        var row = new Row { RowIndex = rowIndex };
        row.Append(cells);
        return row;
    }

    private static Cell TextCell(string reference, string value) => new()
    {
        CellReference = reference,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(value) { Space = SpaceProcessingModeValues.Preserve }),
    };

    private static Cell NumberCell(string reference, string value) => new()
    {
        CellReference = reference,
        DataType = CellValues.Number,
        CellValue = new CellValue(value),
    };

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-import-resilience-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class TextRow
    {
        public string? Name { get; set; }
    }

    private sealed class SparseRow
    {
        public string? Name { get; set; }

        public int Age { get; set; }

        public int Id { get; set; }
    }

    private sealed class StringPair
    {
        public string? Name { get; set; }

        public string? Value { get; set; }
    }

    private sealed class FormulaRow
    {
        public int Amount { get; set; }
    }

    private sealed class WhitespaceRow
    {
        public string? Text { get; set; }

        public int Amount { get; set; }
    }
}
