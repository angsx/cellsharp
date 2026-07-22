using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelDeclarativeValidationTests
{
    private static readonly string[] StatusValues = ["Active", "Inactive", "Pending"];

    [Fact]
    public void AllowedValuesRunsDuringReadAndBeforeCustomValidation()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column.AllowedValues(StatusValues))
            .Column(row => row.Age, column => column
                .Range(18, 120)
                .Validate(age => age % 2 == 0, "Age must be even"))
            .Build();

        try
        {
            Excel.Write(path,
            [
                new ValidationRow { Status = "Active", Age = 18 },
                new ValidationRow { Status = "Unknown", Age = 17 },
            ], schema);

            var result = Excel.Read<ValidationRow>(path, schema);

            Assert.Equal("Active", Assert.Single(result.Rows).Status);
            Assert.Equal(3, result.Errors.Count);
            Assert.Equal("Value must be one of: Active, Inactive, Pending.", result.Errors[0].Message);
            Assert.Equal("Value must be between 18 and 120.", result.Errors[1].Message);
            Assert.Equal("Age must be even", result.Errors[2].Message);
            Assert.All(result.Errors, error => Assert.Equal(ExcelReadErrorCode.ValidationFailed, error.Code));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void NumericAndDateRangesHonorBoundariesAndNullableValues()
    {
        var path = TemporaryPath();
        var minimumDate = new DateTime(2000, 1, 1);
        var maximumDate = new DateTime(2030, 12, 31);
        var schema = Excel.Schema<ValidationRow>()
            .Column(row => row.Age, column => column.Range(18, 120))
            .Column(row => row.BirthDate, column => column.DateBetween(minimumDate, maximumDate))
            .Column(row => row.OptionalScore, column => column.Range(0, 100))
            .Build();

        try
        {
            Excel.Write(path,
            [
                new ValidationRow { Age = 18, BirthDate = minimumDate, OptionalScore = null },
                new ValidationRow { Age = 120, BirthDate = maximumDate, OptionalScore = 100 },
                new ValidationRow { Age = 17, BirthDate = minimumDate.AddDays(-1), OptionalScore = -1 },
                new ValidationRow { Age = 121, BirthDate = maximumDate.AddDays(1), OptionalScore = 101 },
            ], schema);

            var result = Excel.Read<ValidationRow>(path, schema);

            Assert.Equal(2, result.Rows.Count);
            Assert.Equal(6, result.Errors.Count);
            Assert.Contains(result.Errors, error => error.Message == "Value must be between 18 and 120.");
            Assert.Contains(result.Errors, error => error.Message == "Date must be between 2000-01-01 and 2030-12-31.");
            Assert.Contains(result.Errors, error => error.Message == "Value must be between 0 and 100.");
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void DeclarativeValidationConfigurationFailsEarlyForInvalidOrIgnoredColumns()
    {
        Assert.Throws<ArgumentException>(() => Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column.AllowedValues()));
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column.Range("a", "z")));
        Assert.Throws<ArgumentException>(() => Excel.Schema<ValidationRow>()
            .Column(row => row.Age, column => column.Range(120, 18)));
        Assert.Throws<ArgumentException>(() => Excel.Schema<ValidationRow>()
            .Column(row => row.BirthDate, column => column.DateBetween(new DateTime(2030, 1, 1), new DateTime(2000, 1, 1))));
        Assert.Throws<InvalidOperationException>(() => Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column.AllowedValues("Active").Ignore()));
    }

    [Fact]
    public void OptionalDeclarativeColumnIsOnlyValidatedWhenItsHeaderIsPresent()
    {
        var missingPath = TemporaryPath();
        var invalidPath = TemporaryPath();
        var schema = Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column.Optional().AllowedValues(StatusValues))
            .Column(row => row.Age)
            .Build();

        try
        {
            CreateWorkbook(missingPath, ["Age"], ["20"]);
            CreateWorkbook(invalidPath, ["Status", "Age"], ["Unknown", "20"]);

            Assert.True(Excel.Read<ValidationRow>(missingPath, schema).IsValid);
            var invalid = Excel.Read<ValidationRow>(invalidPath, schema);
            Assert.Single(invalid.Errors);
            Assert.Equal("Status", invalid.Errors[0].Header);
        }
        finally
        {
            Delete(missingPath);
            Delete(invalidPath);
        }
    }

    [Fact]
    public void CreateTemplateWritesNativeListNumericAndDateValidation()
    {
        var path = TemporaryPath();
        var minimumDate = new DateTime(2000, 1, 1);
        var maximumDate = new DateTime(2030, 12, 31);
        var schema = Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column
                .Header("Customer Status")
                .Optional()
                .AllowedValues("Active", "In progress, quoted \"value\"", "Active"))
            .Column(row => row.Age, column => column.Range(18, 120))
            .Column(row => row.BirthDate, column => column
                .Format("dd/MM/yyyy")
                .DateBetween(minimumDate, maximumDate))
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema, options => options.FreezeHeaderRow());

            using var document = SpreadsheetDocument.Open(path, false);
            var workbookPart = document.WorkbookPart!;
            var worksheet = MainWorksheetPart(workbookPart).Worksheet!;
            var validations = worksheet.GetFirstChild<DataValidations>()!.Elements<DataValidation>().ToArray();
            var hiddenSheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == "_CellSharpValidation");
            var errors = new OpenXmlValidator().Validate(document).ToArray();

            Assert.Equal(3, validations.Length);
            Assert.Equal(DataValidationValues.List, validations[0].Type!.Value);
            Assert.Equal("A2:A1048576", validations[0].SequenceOfReferences!.InnerText);
            Assert.Equal("'_CellSharpValidation'!$A$1:$A$2", validations[0].GetFirstChild<Formula1>()!.Text);
            Assert.Equal(DataValidationValues.Decimal, validations[1].Type!.Value);
            Assert.Equal("18", validations[1].GetFirstChild<Formula1>()!.Text);
            Assert.Equal("120", validations[1].GetFirstChild<Formula2>()!.Text);
            Assert.Equal(DataValidationValues.Date, validations[2].Type!.Value);
            Assert.Equal("C2:C1048576", validations[2].SequenceOfReferences!.InnerText);
            Assert.Equal(SheetStateValues.Hidden, hiddenSheet.State!.Value);
            Assert.NotNull(worksheet.GetFirstChild<SheetViews>());
            Assert.Empty(errors);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteAppliesNativeValidationWithoutChangingExportedValues()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column.AllowedValues(StatusValues))
            .Column(row => row.Age, column => column.Range(18, 120))
            .Build();

        try
        {
            Excel.Write(path, [new ValidationRow { Status = "Active", Age = 18 }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var rows = MainWorksheetPart(document.WorkbookPart!).Worksheet!.GetFirstChild<SheetData>()!
                .Elements<Row>()
                .ToArray();
            var validations = MainWorksheetPart(document.WorkbookPart!).Worksheet!.GetFirstChild<DataValidations>()!;

            Assert.Equal("Active", Text(rows[1].Elements<Cell>().First()));
            Assert.Equal("18", rows[1].Elements<Cell>().Skip(1).First().CellValue!.Text);
            Assert.Equal(2U, validations.Count!.Value);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void CreateTemplateUsesTheHiddenLookupSheetForLongAllowedValueLists()
    {
        var path = TemporaryPath();
        var values = Enumerable.Range(1, 40)
            .Select(index => $"Status-{index:D2}-{new string('x', 12)}")
            .ToArray();
        var schema = Excel.Schema<ValidationRow>()
            .Column(row => row.Status, column => column.AllowedValues(values))
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var workbookPart = document.WorkbookPart!;
            var validation = MainWorksheetPart(workbookPart).Worksheet!.GetFirstChild<DataValidations>()!
                .Elements<DataValidation>()
                .Single();
            var lookupSheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == "_CellSharpValidation");
            var lookupPart = (WorksheetPart)workbookPart.GetPartById(lookupSheet.Id!);

            Assert.Equal("'_CellSharpValidation'!$A$1:$A$40", validation.GetFirstChild<Formula1>()!.Text);
            Assert.Equal(40, lookupPart.Worksheet!.GetFirstChild<SheetData>()!.Elements<Row>().Count());
        }
        finally
        {
            Delete(path);
        }
    }

    private static void CreateWorkbook(string path, string[] headers, params string[][] values)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData();
        data.AppendChild(Row(headers));
        foreach (var value in values)
        {
            data.AppendChild(Row(value));
        }

        worksheetPart.Worksheet = new Worksheet(data);
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = "Sheet1" });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static Row Row(IEnumerable<string> values)
    {
        var row = new Row();
        foreach (var value in values)
        {
            row.AppendChild(new Cell
            {
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(value)),
            });
        }

        return row;
    }

    private static string Text(Cell cell) => cell.InlineString?.Text?.Text ?? cell.CellValue?.Text ?? string.Empty;

    private static WorksheetPart MainWorksheetPart(WorkbookPart workbookPart)
    {
        var sheet = workbookPart.Workbook!.Sheets!.Elements<Sheet>().Single(sheet => sheet.Name!.Value == "Sheet1");
        return (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-validation-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class ValidationRow
    {
        public string? Status { get; set; }

        public int Age { get; set; }

        public DateTime BirthDate { get; set; }

        public int? OptionalScore { get; set; }
    }
}
