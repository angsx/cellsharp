using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xunit;

namespace CellSharp.Tests;

public sealed class ExcelSheetNameTests
{
    [Fact]
    public void SchemaDefaultsToSheet1AndWritePreservesTheCurrentWorksheetName()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Person>().Column(person => person.Name).Build();

        try
        {
            Excel.Write(path, [new Person { Name = "Ada" }], schema);

            Assert.Equal("Sheet1", schema.SheetName);
            Assert.Equal("Sheet1", SheetName(path));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void WriteAndReadUseTheSchemaWorksheetName()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Person>()
            .SheetName("People")
            .Column(person => person.Name)
            .Build();

        try
        {
            Excel.Write(path, [new Person { Name = "Ada" }], schema);

            Assert.Equal("People", SheetName(path));
            Assert.Equal("Ada", Assert.Single(Excel.Read<Person>(path, schema).Rows).Name);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void CreateTemplateUsesTheSchemaWorksheetName()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Person>()
            .SheetName("People")
            .Column(person => person.Name)
            .Build();

        try
        {
            Excel.CreateTemplate(path, schema);

            Assert.Equal("People", SheetName(path));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void ReadThrowsWhenTheConfiguredWorksheetIsMissing()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Person>()
            .SheetName("People")
            .Column(person => person.Name)
            .Build();

        try
        {
            CreateWorkbook(path, "Overview", ["Name"], ["Ada"]);

            var exception = Assert.Throws<InvalidOperationException>(() => Excel.Read<Person>(path, schema));

            Assert.Contains("Worksheet 'People' was not found", exception.Message);
        }
        finally
        {
            Delete(path);
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData("   ")]
    [InlineData("People/2026")]
    [InlineData("People:2026")]
    [InlineData("People[2026]")]
    [InlineData("People\u0001")]
    public void SheetNameRejectsInvalidExcelWorksheetNames(string sheetName)
    {
        Assert.Throws<ArgumentException>(() => Excel.Schema<Person>().SheetName(sheetName));
    }

    [Fact]
    public void SheetNameRejectsNamesLongerThanExcelAllows()
    {
        var sheetName = new string('A', 32);

        Assert.Throws<ArgumentOutOfRangeException>(() => Excel.Schema<Person>().SheetName(sheetName));
    }

    [Fact]
    public void SheetNamePreservesAllowedCharactersWithoutEscaping()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Person>()
            .SheetName("People's FY26")
            .Column(person => person.Name)
            .Build();

        try
        {
            Excel.Write(path, [new Person { Name = "Ada" }], schema);

            Assert.Equal("People's FY26", SheetName(path));
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void SheetNameWorksWithConvertersAndNativeValidation()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Person>()
            .SheetName("People")
            .Column(person => person.Name, column => column.AllowedValues("Ada", "Grace"))
            .Column(person => person.Status, column => column
                .ConvertWith(StatusConverter.Instance))
            .Build();

        try
        {
            Excel.Write(path, [new Person { Name = "Ada", Status = Status.Active }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var sheet = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>()
                .Single(candidate => candidate.Name!.Value == "People");
            var worksheet = document.WorkbookPart.GetPartById(sheet.Id!) as WorksheetPart;
            var validation = worksheet!.Worksheet!.GetFirstChild<DataValidations>()!
                .Elements<DataValidation>()
                .Single();

            Assert.Equal("People", sheet.Name!.Value);
            Assert.Equal(DataValidationValues.List, validation.Type!.Value);
            Assert.Equal(Status.Active, Assert.Single(Excel.Read<Person>(path, schema).Rows).Status);
        }
        finally
        {
            Delete(path);
        }
    }

    [Fact]
    public void SheetNameDoesNotCollideWithTheInternalValidationWorksheet()
    {
        var path = TemporaryPath();
        var schema = Excel.Schema<Person>()
            .SheetName("_CellSharpValidation")
            .Column(person => person.Name, column => column.AllowedValues("Ada", "Grace"))
            .Build();

        try
        {
            Excel.Write(path, [new Person { Name = "Ada" }], schema);

            using var document = SpreadsheetDocument.Open(path, false);
            var sheets = document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().ToArray();

            Assert.Contains(sheets, sheet => sheet.Name!.Value == "_CellSharpValidation" && sheet.State is null);
            Assert.Contains(sheets, sheet => sheet.Name!.Value == "_CellSharpValidation2" && sheet.State!.Value == SheetStateValues.Hidden);
            Assert.Equal("Ada", Assert.Single(Excel.Read<Person>(path, schema).Rows).Name);
        }
        finally
        {
            Delete(path);
        }
    }

    private static string SheetName(string path)
    {
        using var document = SpreadsheetDocument.Open(path, false);
        return document.WorkbookPart!.Workbook!.Sheets!.Elements<Sheet>().Single().Name!.Value!;
    }

    private static void CreateWorkbook(string path, string sheetName, params string[][] rows)
    {
        using var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook);
        var workbookPart = document.AddWorkbookPart();
        workbookPart.Workbook = new Workbook();
        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        var data = new SheetData();
        foreach (var values in rows)
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

            data.AppendChild(row);
        }

        worksheetPart.Worksheet = new Worksheet(data);
        var sheets = workbookPart.Workbook!.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1U, Name = sheetName });
        worksheetPart.Worksheet!.Save();
        workbookPart.Workbook!.Save();
    }

    private static string TemporaryPath() => Path.Combine(Path.GetTempPath(), $"cellsharp-sheet-name-{Guid.NewGuid():N}.xlsx");

    private static void Delete(string path)
    {
        if (File.Exists(path))
        {
            File.Delete(path);
        }
    }

    private sealed class Person
    {
        public string? Name { get; set; }

        public Status? Status { get; set; }
    }

    private sealed class Status : IEquatable<Status>
    {
        internal static Status Active { get; } = new("ACTIVE");
        internal static Status Inactive { get; } = new("INACTIVE");

        internal Status(string code) => Code = code;

        internal string Code { get; }

        public bool Equals(Status? other) => other is not null && Code == other.Code;

        public override bool Equals(object? obj) => Equals(obj as Status);

        public override int GetHashCode() => Code.GetHashCode();
    }

    private sealed class StatusConverter : IExcelValueConverter<Status?, string>
    {
        internal static StatusConverter Instance { get; } = new();

        public string Write(Status? value) => value?.Code ?? string.Empty;

        public bool TryRead(string value, out Status? converted)
        {
            converted = value == "ACTIVE" ? Status.Active : value == "INACTIVE" ? Status.Inactive : null;
            return converted is not null;
        }
    }
}
